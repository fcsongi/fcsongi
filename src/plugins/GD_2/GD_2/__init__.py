"""
This is where the implementation of the plugin code goes.
The GD_1-class is imported from both run_plugin.py and run_debug.py
"""

import json
import sys
import logging
from webgme_bindings import PluginBase

# Setting up a logger
logger = logging.getLogger('GD_2')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

libNode = None

def search_in_library(self, core, target):                              #meta_type(string): {Device, Country, Vendor, Region} | target(string) is the node's name we're searching
    childrens = core.load_children(self.META["Countries"])                      #loading all children from library
    for child in childrens:
        node = core.load_by_path(self.root_node, child['nodePath'])                  #loading the node by path
        get_child(self, core, node, target)


def get_child(self, core, node, target):
    global libNode
    # logger.info(" Device: '{0}'".format(core.get_attribute(node,'name')))        #logging the found child's name
    children = core.load_children(node)                                         #getting all the children of a node
    for child in children:
        if core.get_attribute(child, 'name') == target:                    #if we found the searched country's node
            # logger.info("Found node: '{0}'".format(core.get_attribute(child, 'name')))
            libNode = child
            break                                                        #return it
        else:                                                               #else
            get_child(self, core, child, target)                                     #search in the child's children


class GD_2(PluginBase):
    def main(self):
        global libNode
        core = self.core
        root_node = self.root_node
        active_node = self.active_node

        with open("./imports/costumRegions.json") as f:     #loading regions json file into data list
            regions = json.load(f)

        with open("./imports/countries2.json") as f:    #loading countries json file into data list
            countries = json.load(f)

        name = core.get_attribute(active_node, 'name')   #get the name of active_node
        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))   #logging the active_node's name

        #getting the meta of target file which we want to create
        region_meta = self.META["CustomRegion"]

        i = 1  #temporary variable that helps to push the new child off from the old child so they does not stack
        for region in regions:
            child1 = core.create_child(active_node, region_meta)                #creating a child node of the active node
            position_item = core.get_registry(active_node,'position')           #getting position of the active node
            position_item['y'] = position_item['y'] + 50 * i                    #changing the position variable
            position_item['x'] += 400
            core.set_registry(child1, 'position', position_item)                #changing the child's position
            core.set_attribute(child1, 'name', region['name'])                  #changing the child's name
            i+=1                                                                #update the temporary variable

            for country in countries:
                if country["region"] == core.get_attribute(child1, 'name'):     #check if the selected country is in it's region
                    search_in_library(self,core,country["country"])
                    core.add_member(child1, "MemberCountries", libNode)

        # creating a commit for the update
        commit_info = self.util.save(root_node, self.commit_hash, 'test', 'Python plugin created regions with countries')
        logger.info('commited:{0}'.format(commit_info))
