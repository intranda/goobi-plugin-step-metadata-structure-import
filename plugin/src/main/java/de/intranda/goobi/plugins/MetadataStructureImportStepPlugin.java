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
import org.goobi.production.plugin.interfaces.IStepPluginVersion2;

import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.VariableReplacer;
import de.sub.goobi.helper.exceptions.SwapException;
import lombok.Getter;
import lombok.extern.log4j.Log4j2;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.Fileformat;
import ugh.exceptions.UGHException;

@PluginImplementation
@Log4j2
public class MetadataStructureImportStepPlugin implements IStepPluginVersion2 {

    private static final long serialVersionUID = -2724211643330484400L;

    @Getter
    private String title = "intranda_step_MetadataStructureImport";
    @Getter
    private Step step;

    private Process process;

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

    @Override
    public void initialize(Step step, String returnPath) {
        this.step = step;
        process = step.getProzess();
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

        identifierColumnName = config.getString("/identifierColumnName");
        doctypeColumnName = config.getString("/doctypeColumnName");
        hierarchyColumnName = config.getString("/hierarchyColumnName");
        imageStartColumnName = config.getString("/imageStartColumnName");
        imageEndColumnName = config.getString("/imageEndColumnName");
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

        List<DocStruct> children = logical.getAllChildren();
        if (children != null) {
            for (DocStruct child : children) {
                logical.removeChild(child);
            }
        }

        // TODO if physical is empty, generate pagination
        if (physical.getAllChildren() == null) {

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

        Map<String, Integer> headerOrder = new HashMap<>();

        // open excel file
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

            //  read and validate the header row
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
                        case ERROR:
                        case BLANK:
                        case _NONE:
                        default:
                            value = "";
                            break;
                    }
                    headerOrder.put(value, i);
                }
            }

            // find out the first data row
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

                String identifier = getCellValue(row, headerOrder.get(identifierColumnName));

                int hierarchy = Integer.parseInt(getCellValue(row, headerOrder.get(hierarchyColumnName)));

                DocStruct currentType = null;

                // TODO find correct position
                // if current element hierarchy is higher than last element, its a child element
                // if it has the same number, its a sibling
                // if it is smaller, go upwards to find the right parent element, insert as last

                // TODO get opac record for identifier

                // copy metadata to the new docstruct

                // TODO get additional metadata from excel document
                // overwrite/insert new metadata
                for (Column col : columns) {
                    int colId = headerOrder.get(col.getColumnName());
                    String colVal = getCellValue(row, colId);

                }
            }

        } catch (IOException e) {
            log.error(e);
        }

        boolean successful = true;
        // your logic goes here

        log.info("MetadataStructureImport step plugin executed");
        if (!successful) {
            return PluginReturnValue.ERROR;
        }
        return PluginReturnValue.FINISH;
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
                value = String.valueOf(cell.getNumericCellValue());
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
