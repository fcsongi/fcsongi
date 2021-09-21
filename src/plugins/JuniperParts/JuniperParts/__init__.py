"""
This is where the implementation of the plugin code goes.
The JuniperParts-class is imported from both run_plugin.py and run_debug.py
"""
import sys
import logging
import pandas
import json
from webgme_bindings import PluginBase

# Setup a logger
logger = logging.getLogger('JuniperParts')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class JuniperParts(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node

        name = core.get_attribute(active_node, 'name')

        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))

        #loads excel as json
        excel_df = pandas.read_excel('./imports/juniperSupport.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        supports = json.loads(result)

        #loads excel as json
        excel_df = pandas.read_excel('./imports/juniperSoftwares.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        softwares = json.loads(result)

        #loads excel as json
        excel_df = pandas.read_excel('./imports/juniperLicenses.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        licenses = json.loads(result)

        excel_df = pandas.read_excel('./imports/juniperAccessories.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        accessories = json.loads(result)

        bundles = core.load_children(active_node)

        for bundle in bundles:
            i = 1
            for row in supports:
                if core.get_attribute(bundle,'name') in row['code']:
                    child_node = core.create_child(bundle, self.META["Part"])
                    position_item = core.get_registry(active_node,'position')           
                    position_item['y'] = position_item['y'] + 50 * i                   
                    core.set_registry(child_node, 'position', position_item)                
                    core.set_attribute(child_node, 'type', "Support")             
                    core.set_attribute(child_node, 'name', row['code'])
                    core.set_attribute(child_node, 'vendorCode', row['vendorCode'])
                    core.set_attribute(child_node, 'cost', row['cost'])
                    core.set_attribute(child_node, 'discount', row['discount'])  
                    core.set_attribute(child_node, 'description', row['description'])                
                    i+=1

        for bundle in bundles:
            i = 1
            for row in softwares:
                if core.get_attribute(bundle,'name') in row['code']:
                    child_node = core.create_child(bundle, self.META["Part"])
                    position_item = core.get_registry(active_node,'position')           
                    position_item['y'] = position_item['y'] + 50 * i 
                    position_item['x'] = position_item['x'] + 250                  
                    core.set_registry(child_node, 'position', position_item)
                    core.set_attribute(child_node, 'type', "Software")             
                    core.set_attribute(child_node, 'name', row['code'])
                    core.set_attribute(child_node, 'vendorCode', row['vendorCode'])
                    core.set_attribute(child_node, 'cost', row['cost'])
                    core.set_attribute(child_node, 'discount', row['discount'])
                    core.set_attribute(child_node, 'description', row['description'])
                    i+=1
        
        for bundle in bundles:
            i = 1
            for row in licenses:
                if core.get_attribute(bundle,'name') in row['code']:
                    child_node = core.create_child(bundle, self.META["Part"])
                    position_item = core.get_registry(active_node,'position')           
                    position_item['y'] = position_item['y'] + 50 * i 
                    position_item['x'] = position_item['x'] + 500                  
                    core.set_registry(child_node, 'position', position_item)                
                    core.set_attribute(child_node, 'type', "License")             
                    core.set_attribute(child_node, 'name', row['code'])
                    core.set_attribute(child_node, 'vendorCode', row['vendorCode'])
                    core.set_attribute(child_node, 'cost', row['cost'])
                    core.set_attribute(child_node, 'discount', row['discount'])
                    core.set_attribute(child_node, 'description', row['description'])                
                    i+=1
        
        for bundle in bundles:
            i = 1
            for row in accessories:
                if core.get_attribute(bundle,'name') in row['compatibility']:
                    child_node = core.create_child(bundle, self.META["Part"])
                    position_item = core.get_registry(active_node, 'position')
                    position_item['y'] = position_item['y'] + 50 * i
                    position_item['x'] = position_item['x'] + 750                       
                    core.set_registry(child_node, 'position', position_item)
                    core.set_attribute(child_node, 'type', 'Accessory')
                    core.set_attribute(child_node, 'name', row['code'])
                    core.set_attribute(child_node, 'vendorCode', row['vendorCode'])
                    core.set_attribute(child_node, 'description', row['description'])
                    core.set_attribute(child_node, 'cost', row['cost'])
                    core.set_attribute(child_node, 'discount', row['discount'])
                    i+=1

        commit_info = self.util.save(root_node, self.commit_hash, 'test', 'Python plugin populated the Juniper vendor bundles with parts')
        logger.info('committed :{0}'.format(commit_info))
