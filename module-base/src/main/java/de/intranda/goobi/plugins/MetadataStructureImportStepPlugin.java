package de.intranda.goobi.plugins;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;

/**
 * This file is part of a plugin for Goobi - a Workflow tool for the support of mass digitization.
 *
 * Visit the websites for more information.
 *          - https://goobi.io
 *          - https://www.intranda.com
 *          - https://github.com/intranda/goobi
 *
 * This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free
 * Software Foundation; either version 2 of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 * FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59
 * Temple Place, Suite 330, Boston, MA 02111-1307 USA
 *
 */

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.io.ByteOrderMark;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.goobi.beans.Process;
import org.goobi.beans.Step;
import org.goobi.production.enums.PluginGuiType;
import org.goobi.production.enums.PluginReturnValue;
import org.goobi.production.enums.PluginType;
import org.goobi.production.enums.StepReturnValue;
import org.goobi.production.plugin.interfaces.IOpacPlugin;
import org.goobi.production.plugin.interfaces.IStepPluginVersion2;

import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.VariableReplacer;
import de.sub.goobi.helper.exceptions.DAOException;
import de.sub.goobi.helper.exceptions.SwapException;
import de.sub.goobi.metadaten.MetadatenImagesHelper;
import de.unigoettingen.sub.search.opac.ConfigOpac;
import de.unigoettingen.sub.search.opac.ConfigOpacCatalogue;
import lombok.Getter;
import lombok.extern.log4j.Log4j2;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.Corporate;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.dl.MetadataType;
import ugh.dl.Person;
import ugh.dl.Prefs;
import ugh.dl.Reference;
import ugh.exceptions.MetadataTypeNotAllowedException;
import ugh.exceptions.PreferencesException;
import ugh.exceptions.TypeNotAllowedForParentException;
import ugh.exceptions.UGHException;
import ugh.exceptions.WriteException;

@PluginImplementation
@Log4j2
public class MetadataStructureImportStepPlugin implements IStepPluginVersion2 {

    private static final long serialVersionUID = -2724211643330484400L;

    @Getter
    private String title = "intranda_step_MetadataStructureImport";
    @Getter
    private Step step;

    private Process process;

    private Prefs prefs;

    private String excelFolder;

    private int headerRowNumber;
    private int dataRowNumber;
    private int lastDataRow;
    private transient List<Column> columns;

    private String identifierColumnName;
    private String doctypeColumnName;
    private String hierarchyColumnName;
    private String imageStartColumnName;
    private String imageEndColumnName;

    private String opacName;
    private String opacSearchField;

    private Map<String, String> docstructs;

    @Override
    public void initialize(Step step, String returnPath) {
        this.step = step;
        process = step.getProzess();
        prefs = process.getRegelsatz().getPreferences();

        // load configuration
        SubnodeConfiguration config = ConfigPlugins.getProjectAndStepConfig(title, step);

        excelFolder = config.getString("/excelFolder");

        headerRowNumber = config.getInt("/rowHeader", 1);
        dataRowNumber = config.getInt("/rowDataStart", 2);
        lastDataRow = config.getInt("/rowDataEnd", 99999);
        columns = new ArrayList<>();

        List<HierarchicalConfiguration> hcl = config.configurationsAt("/column");
        for (HierarchicalConfiguration hc : hcl) {
            Column col = new Column();
            col.setColumnName(hc.getString("@columnName"));
            col.setMetadataName(hc.getString("@metadata", ""));
            columns.add(col);
        }

        docstructs = new HashMap<>();
        hcl = config.configurationsAt("/docstruct");
        for (HierarchicalConfiguration hc : hcl) {
            docstructs.put(hc.getString("@label"), hc.getString("@value"));
        }

        identifierColumnName = config.getString("/identifierColumnName");
        doctypeColumnName = config.getString("/doctypeColumnName");
        hierarchyColumnName = config.getString("/hierarchyColumnName");
        imageStartColumnName = config.getString("/imageStartColumnName");
        imageEndColumnName = config.getString("/imageEndColumnName");

        opacName = config.getString("/opacName");
        opacSearchField = config.getString("/searchField");
    }

    @Override
    public PluginReturnValue run() {
        // open metadata file

        Fileformat fileformat = null;
        DigitalDocument digDoc = null;
        try {
            fileformat = process.readMetadataFile();
            digDoc = fileformat.getDigitalDocument();
        } catch (UGHException | IOException | SwapException e) {
            log.error(e);
            // cannot read metadata file, abort.
            return PluginReturnValue.ERROR;

        }
        DocStruct logical = digDoc.getLogicalDocStruct();
        DocStruct physical = digDoc.getPhysicalDocStruct();

        VariableReplacer replacer = new VariableReplacer(digDoc, process.getRegelsatz().getPreferences(), process, step);

        // clear metadata file, remove existing structure elements
        if (logical.getAllChildren() != null) {
            List<DocStruct> children = new ArrayList<>(logical.getAllChildren());
            if (children != null) {
                for (DocStruct child : children) {
                    List<Reference> refs = new ArrayList<>(child.getAllToReferences());
                    for (ugh.dl.Reference ref : refs) {
                        child.removeReferenceTo(ref.getTarget());
                    }
                    logical.removeChild(child);
                }
            }
        }

        // create pagination, if missing
        if (physical.getAllChildren() == null) {
            MetadatenImagesHelper imagehelper = new MetadatenImagesHelper(prefs, digDoc);
            try {
                imagehelper.createPagination(process, process.getImagesTifDirectory(true));
            } catch (TypeNotAllowedForParentException | IOException | SwapException | DAOException e) {
                log.error(e);
            }
        }
        List<DocStruct> pages = physical.getAllChildren();

        // load configured opac catalogue
        IOpacPlugin myImportOpac = null;
        ConfigOpacCatalogue coc = null;
        if (StringUtils.isNotBlank(opacName)) {
            for (ConfigOpacCatalogue configOpacCatalogue : ConfigOpac.getInstance().getAllCatalogues("")) {
                if (configOpacCatalogue.getTitle().equals(opacName)) {
                    myImportOpac = configOpacCatalogue.getOpacPlugin();
                    coc = configOpacCatalogue;
                }
            }
        }

        // find excel file in configured folder
        Path excelFile = null;
        Path path = Paths.get(replacer.replace(excelFolder));
        if (!StorageProvider.getInstance().isDirectory(path)) {
            // excel folder not found, abort
            return PluginReturnValue.ERROR;
        }

        List<Path> dataInFolder = StorageProvider.getInstance().listFiles(path.toString());
        for (Path p : dataInFolder) {
            if (p.getFileName().toString().endsWith("xlsx")) {
                excelFile = p;
            }
        }
        if (excelFile == null) {
            // excel file not found, abort
            return PluginReturnValue.ERROR;
        }

        // open excel file
        Map<String, Integer> headerOrder = new HashMap<>();
        try (InputStream fileInputStream = StorageProvider.getInstance().newInputStream(excelFile);
                BOMInputStream in = BOMInputStream.builder()
                        .setPath(excelFile)
                        .setByteOrderMarks(ByteOrderMark.UTF_8)
                        .setInclude(false)
                        .get();
                Workbook wb = WorkbookFactory.create(in)) {
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();

            int rowCounter = 0;

            //  find the header row
            Row headerRow = null;
            while (rowCounter < headerRowNumber) {
                headerRow = rowIterator.next();
                rowCounter++;
            }

            //  read the header row
            int numberOfCells = headerRow.getLastCellNum();
            for (int i = 0; i < numberOfCells; i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    String value = null;
                    switch (cell.getCellType()) {
                        case BOOLEAN:
                            value = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            value = cell.getCellFormula();
                            break;
                        case NUMERIC:
                            value = String.valueOf(cell.getNumericCellValue());
                            break;
                        case STRING:
                            value = cell.getStringCellValue();
                            break;
                        case ERROR, BLANK, _NONE:
                        default:
                            value = "";
                            break;
                    }
                    headerOrder.put(value, i);
                }
            }

            // find the first data row
            while (rowCounter < dataRowNumber - 1) {
                headerRow = rowIterator.next();
                rowCounter++;
            }

            DocStruct lastElement = logical;
            int lastHierarchy = 0;

            // run through all the data rows
            while (rowIterator.hasNext() && rowCounter < lastDataRow) {

                // for each line in excel file:

                // generate structure element
                // parent element is the last element with smaller hierarchy level (or the root element)
                // add metadata from configured columns
                // create page assignments based on excel data
                // opac request if configured and identifier is known
                // excel data has higher priority than opac data

                Row row = rowIterator.next();
                rowCounter++;
                int lastColumn = row.getLastCellNum();
                if (lastColumn == -1) {
                    continue;
                }

                String docType = getCellValue(row, headerOrder.get(doctypeColumnName));
                int hierarchy = Integer.parseInt(getCellValue(row, headerOrder.get(hierarchyColumnName)));

                String identifier = getCellValue(row, headerOrder.get(identifierColumnName));

                int startPageNo = Integer.parseInt(getCellValue(row, headerOrder.get(imageStartColumnName)));
                int endPageNo = Integer.parseInt(getCellValue(row, headerOrder.get(imageEndColumnName)));

                DocStruct currentDocStruct = digDoc.createDocStruct(prefs.getDocStrctTypeByName(docstructs.get(docType)));

                // skip first element as it is the publication type itself
                if (hierarchy != 0) {

                    // if current element hierarchy is higher than last element, its a child element of the last element
                    if (hierarchy > lastHierarchy) {
                        lastElement.addChild(currentDocStruct);
                    }
                    // if it has the same number, its a sibling, add it as child element of the parent
                    else if (hierarchy == lastHierarchy) {
                        lastElement.getParent().addChild(currentDocStruct);
                    } else {
                        // if it is smaller, go upwards to find the right parent element, insert as last
                        while (hierarchy < lastHierarchy) {
                            lastElement = lastElement.getParent();
                            lastHierarchy--;
                        }
                        lastElement.getParent().addChild(currentDocStruct);
                    }

                    lastElement = currentDocStruct;
                    lastHierarchy = hierarchy;

                    //  get opac record for identifier

                    if (StringUtils.isNotBlank(identifier) && coc != null && myImportOpac != null) {
                        getOpacRequest(currentDocStruct, myImportOpac, coc, identifier);
                    }

                    // copy metadata from response to the new docstruct

                    // assign pages
                    List<DocStruct> pagesToAssign = pages.subList(startPageNo - 1, endPageNo);

                    for (DocStruct page : pagesToAssign) {
                        currentDocStruct.addReferenceTo(page, "logical_physical");
                    }

                    // get additional metadata from excel document
                    for (Column col : columns) {
                        int colId = headerOrder.get(col.getColumnName());
                        String colVal = getCellValue(row, colId);

                        // overwrite/insert new metadata
                        MetadataType metadataType = prefs.getMetadataTypeByName(col.getMetadataName());

                        List<? extends Metadata> metadataList = currentDocStruct.getAllMetadataByType(metadataType);
                        if (!metadataList.isEmpty()) {
                            Metadata metadata = metadataList.get(0);
                            metadata.setValue(colVal);
                        } else {
                            Metadata metadata = new Metadata(metadataType);
                            metadata.setValue(colVal);
                            currentDocStruct.addMetadata(metadata);
                        }
                    }
                }
            }
        } catch (IOException | UGHException e) {
            log.error(e);
        }

        try {
            process.writeMetadataFile(fileformat);
        } catch (WriteException | PreferencesException | IOException | SwapException e) {
            log.error(e);
        }

        return PluginReturnValue.FINISH;
    }

    private void getOpacRequest(DocStruct currentDocstruct, IOpacPlugin myImportOpac, ConfigOpacCatalogue coc, String identifier)
            throws PreferencesException {
        Fileformat opacResponse = null;
        try {
            opacResponse = myImportOpac.search(opacSearchField, identifier, coc, prefs);
        } catch (Exception e) {
            log.error(e);
        }
        if (opacResponse != null) {
            DocStruct opacLogical = opacResponse.getDigitalDocument().getLogicalDocStruct();
            if (opacLogical.getType().isAnchor()) {
                opacLogical = opacLogical.getAllChildren().get(0);
            }
            if (opacLogical.getAllMetadata() != null) {
                for (Metadata md : opacLogical.getAllMetadata()) {
                    try {
                        Metadata copy = new Metadata(md.getType());
                        copy.setValue(md.getValue());
                        copy.setAutorityFile(md.getAuthorityID(), md.getAuthorityURI(), md.getAuthorityValue());
                        currentDocstruct.addMetadata(copy);
                    } catch (MetadataTypeNotAllowedException e) {
                        log.debug(e);
                    }
                }
            }
            if (opacLogical.getAllPersons() != null) {
                for (Person p : opacLogical.getAllPersons()) {
                    try {
                        Person copy = new Person(p.getType());
                        copy.setFirstname(p.getFirstname());
                        copy.setLastname(p.getLastname());
                        copy.setAutorityFile(p.getAuthorityID(), p.getAuthorityURI(), p.getAuthorityValue());
                        currentDocstruct.addPerson(copy);
                    } catch (MetadataTypeNotAllowedException e) {
                        log.debug(e);
                    }

                }
            }
            if (opacLogical.getAllCorporates() != null) {
                for (Corporate c : opacLogical.getAllCorporates()) {
                    try {
                        Corporate copy = new Corporate(c.getType());
                        copy.setMainName(c.getMainName());
                        copy.setAutorityFile(c.getAuthorityID(), c.getAuthorityURI(), c.getAuthorityValue());
                        currentDocstruct.addCorporate(copy);
                    } catch (MetadataTypeNotAllowedException e) {
                        log.debug(e);

                    }

                }
            }
        }
    }

    @Override
    public PluginGuiType getPluginGuiType() {
        return PluginGuiType.NONE;
    }

    @Override
    public String getPagePath() {
        return "/uii/plugin_step_MetadataStructureImport.xhtml";
    }

    @Override
    public PluginType getType() {
        return PluginType.Step;
    }

    @Override
    public String cancel() {
        return "";
    }

    @Override
    public String finish() {
        return "";
    }

    @Override
    public int getInterfaceVersion() {
        return 0;
    }

    @Override
    public HashMap<String, StepReturnValue> validate() {
        return null; //NOSONAR
    }

    @Override
    public boolean execute() {
        PluginReturnValue ret = run();
        return ret != PluginReturnValue.ERROR;
    }

    public String getCellValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        String value = "";
        switch (cell.getCellType()) {
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                value = String.valueOf((long) cell.getNumericCellValue());
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            default:
                // none, error, blank
                value = "";
                break;
        }
        return value;
    }

}
