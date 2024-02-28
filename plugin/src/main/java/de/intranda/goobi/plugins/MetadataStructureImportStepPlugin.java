package de.intranda.goobi.plugins;

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

import org.apache.commons.configuration.SubnodeConfiguration;
import org.goobi.beans.Step;
import org.goobi.production.enums.PluginGuiType;
import org.goobi.production.enums.PluginReturnValue;
import org.goobi.production.enums.PluginType;
import org.goobi.production.enums.StepReturnValue;
import org.goobi.production.plugin.interfaces.IStepPluginVersion2;

import de.sub.goobi.config.ConfigPlugins;
import lombok.Getter;
import lombok.extern.log4j.Log4j2;
import net.xeoh.plugins.base.annotations.PluginImplementation;

@PluginImplementation
@Log4j2
public class MetadataStructureImportStepPlugin implements IStepPluginVersion2 {

    private static final long serialVersionUID = -2724211643330484400L;

    @Getter
    private String title = "intranda_step_MetadataStructureImport";
    @Getter
    private Step step;

    @Override
    public void initialize(Step step, String returnPath) {
        this.step = step;

        // load configuration
        SubnodeConfiguration myconfig = ConfigPlugins.getProjectAndStepConfig(title, step);
    }

    @Override
    public PluginReturnValue run() {
        // find excel file

        // open excel file

        // open metadata file

        // clear metadata file, remove existing structure elements

        // for each line in excel file:

        // generate structure element
        // parent element is the last element with smaller hierarchy level (or the root element)
        // add metadata from configured columns
        // create page assignments based on excel data
        // opac request if configured and identifier is known
        // excel data has higher priority than opac data

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

}
