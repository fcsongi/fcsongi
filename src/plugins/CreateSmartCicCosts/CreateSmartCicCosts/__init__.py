"""
This is where the implementation of the plugin code goes.
The CreateSmartCicCosts-class is imported from both run_plugin.py and run_debug.py
"""
import pandas
import json
import sys
import logging
from webgme_bindings import PluginBase

# Setup a logger
logger = logging.getLogger('CreateSmartCicCosts')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

#searches custom region nodes
def searchCustomRegions(self, core, active_node, data):
    # logger.info(" Device: '{0}'".format(core.get_attribute(node,'name')))        #logging the found child's name
    children = core.load_children(self.META["Countries"])                                         #getting all the children of a node
    for child in children:
        if core.get_meta_type(child) == self.META["CustomRegion"]:                    #if we found the searched country's node
            # logger.info("Found node: '{0}'".format(core.get_attribute(child, 'name')))
                        #changing the child's position
            createPricingConnection(self, core, active_node, child, data)
  

def createPricingConnection(self, core, active_node, region, data):                              #meta_type(string): {Device, Country, Vendor, Region} | target(string) is the node's name we're searching
    childrens = core.load_children(self.META["Vendors"])
    for child in childrens:
        if core.get_meta_type(child) == self.META["CustomEquipmentGroup"]:
            connection = core.create_child(active_node,self.META["Pricing"])
            core.set_pointer(connection,'src',child)
            core.set_pointer(connection,'dst',region)

            for row in data:
                if row["Region"] == core.get_attribute(region,'name'):
                    core.set_attribute(connection, 'installCost', row['Installation cost'])
                    core.set_attribute(connection, 'additionalInstallCost', row['Additional hour install cost'])
                    core.set_attribute(connection, 'bronzeCost', row['Bronze cost'])
                    core.set_attribute(connection, 'additionalBronzeCost', row['Additional hour Bronze cost'])
                    core.set_attribute(connection, 'silverCost', row['Silver cost'])
                    core.set_attribute(connection, 'additionalSilverCost', row['Additional hour Silver cost'])
                    core.set_attribute(connection, 'goldCost', row['Gold cost'])
                    core.set_attribute(connection, 'additionalGoldCost', row['Additional hour Gold cost'])



class CreateSmartCicCosts(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node

        name = core.get_attribute(active_node, 'name')

        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))
                
        excel_df = pandas.read_excel('./imports/SmartCicPrices.xlsx', sheet_name='SmartCIC',dtype=object)
        result = excel_df.to_json(orient="records")
        data = json.loads(result)

        searchCustomRegions(self, core, active_node, data)

        commit_info = self.util.save(root_node, self.commit_hash, 'test', 'Python plugin updated the model')
        logger.info('committed :{0}'.format(commit_info))
