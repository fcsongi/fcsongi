"""
This is where the implementation of the plugin code goes.
The GD_1-class is imported from both run_plugin.py and run_debug.py
"""

import pandas
import json
import sys
import logging
from webgme_bindings import PluginBase

# Setting up a logger
logger = logging.getLogger('GD_1')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class GD_1(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node

        with open("./imports/regions.json") as f:     #loading regions json file into data list
            regions = json.load(f)

        excel_df = pandas.read_excel('./imports/countries.xlsx', na_filter=False)
        result = excel_df.to_json(orient="records")
        countries = json.loads(result)

        name = core.get_attribute(active_node, 'name')   #get the name of active_node
        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))   #logging the active_node's name

        #getting the meta of target file which we want to create
        region_meta = self.META["Region"]
        country_meta = self.META["Country"]

        i = 1  #temporary variable that helps to push the new child off from the old child so they does not stack
        for region in regions:
            child1 = core.create_child(active_node, region_meta)                #creating a child node of the active node
            position_item = core.get_registry(active_node,'position')           #getting position of the active node
            position_item['y'] = position_item['y'] + 50 * i                    #changing the position variable
            core.set_registry(child1, 'position', position_item)                #changing the child's position
            core.set_attribute(child1, 'name', region['name'])                  #changing the child's name
            i+=1                                                                #update the temporary variable

            j = 1   #temporary variable that helps to push the new child off from the old child so they does not stack
            for country in countries:
                if country["region"] == core.get_attribute(child1, 'name'):     #check if the selected country is in it's region
                    child2 = core.create_child(child1, country_meta)            #creating a child node of the active node
                    position_item = core.get_registry(active_node,'position')   #getting position of the active node
                    position_item['y'] = position_item['y'] + 50 * j            #changing the position variable
                    core.set_registry(child2, 'position', position_item)        #changing the child's position
                    core.set_attribute(child2, 'name', country["Country"])         #changing the child's name
                    if country["Country"] == "Namibia":
                        country["2code"] = "NA"
                    core.set_attribute(child2, 'isoCode2', country["2code"])    #changing the child's isoCode2
                    core.set_attribute(child2, 'isoCode3', country["3code"])    #changing the child's isoCode3
                    core.set_attribute(child2, 'deliveryCategory', country["category"])        #changing the child's category
                    if country["category"] == "C":
                        core.set_attribute(child2, 'MRC_license', country["MRC license"])
                    core.set_attribute(child2, 'DHL code', country["DHL code"])        #changing the child's DHLcode
                    core.set_attribute(child2, 'EU/Not EU', country["EU/Not EU"])
                    j+=1    #update the temporary variable

        #creating a commit for the update
        commit_info = self.util.save(root_node, self.commit_hash, 'test', 'Python plugin created regions with countries')
        logger.info('commited:{0}'.format(commit_info))
