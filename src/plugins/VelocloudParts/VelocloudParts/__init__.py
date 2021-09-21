"""
This is where the implementation of the plugin code goes.
The VelocloudParts-class is imported from both run_plugin.py and run_debug.py
"""
import sys
import logging
import pandas
import json
from webgme_bindings import PluginBase

# Setup a logger
logger = logging.getLogger('VelocloudParts')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class VelocloudParts(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node

        name = core.get_attribute(active_node, 'name')

        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))

        #loads excel as json
        excel_df = pandas.read_excel('./imports/velocloudSupport.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        supports = json.loads(result)

        #loads excel as json
        excel_df = pandas.read_excel('./imports/velocloudSoftwares.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        softwares = json.loads(result)

        #loads excel as json
        excel_df = pandas.read_excel('./imports/velocloudLicenses.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        licenses = json.loads(result)

        excel_df = pandas.read_excel('./imports/velocloudAccessories.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        accessories = json.loads(result)

        bundles = core.load_children(active_node)

        for bundle in bundles:
            i = 1
            for row in supports:
                setup_code = str(row['code']).split("-")
                if core.get_attribute(bundle,'name') == setup_code[0]:
                    child_node = core.create_child(bundle, self.META['Part'])
                    position_item = core.get_registry(active_node,'position')           
                    position_item['y'] = position_item['y'] + 50 * i
                    core.set_registry(child_node, 'position', position_item)                
                    core.set_attribute(child_node, 'type', "Support")
                    core.set_attribute(child_node, 'name', row['code'])
                    core.set_attribute(child_node, 'vendorCode', row['vendorCode'])
                    core.set_attribute(child_node, 'cost', row['cost'])
                    core.set_attribute(child_node, 'discount', row['discount'])
                    core.set_attribute(child_node, 'upgradeMargin', row['upgradeMargin'])
                    core.set_attribute(child_node, 'description', row['description'])
                    core.set_attribute(child_node, 'upgradeDescription', row['upgradeDescription'])      
                    i+=1

        for bundle in bundles:
            i = 1
            for row in softwares:
                min_bandwidth = core.get_attribute(bundle,'min_bandwidth')
                max_bandwidth = core.get_attribute(bundle,'max_bandwidth')
                if "M" in str(min_bandwidth).split(" ")[1]:
                    min_bandwidth = int(str(min_bandwidth).split(" ")[0])
                else:
                    min_bandwidth = int(str(min_bandwidth).split(" ")[0]) * 1000
                if "M" in str(max_bandwidth).split(" ")[1]:
                    max_bandwidth = int(str(max_bandwidth).split(" ")[0])
                else:
                    max_bandwidth = int(str(max_bandwidth).split(" ")[0]) * 1000
                if row['bandwidth'][3] == "M":
                    bandwidth = int(row['bandwidth'][:3])
                else:
                    bandwidth = int(row['bandwidth'][:3]) * 1000
                
                if bandwidth >= min_bandwidth and bandwidth <= max_bandwidth:
                    child_node = core.create_child(bundle, self.META['Part'])
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
                min_bandwidth = core.get_attribute(bundle,'min_bandwidth')
                max_bandwidth = core.get_attribute(bundle,'max_bandwidth')
                if "M" in str(min_bandwidth).split(" ")[1]:
                    min_bandwidth = int(str(min_bandwidth).split(" ")[0])
                else:
                    min_bandwidth = int(str(min_bandwidth).split(" ")[0]) * 1000
                if "M" in str(max_bandwidth).split(" ")[1]:
                    max_bandwidth = int(str(max_bandwidth).split(" ")[0])
                else:
                    max_bandwidth = int(str(max_bandwidth).split(" ")[0]) * 1000
                if row['bandwidth'][3] == "M":
                    bandwidth = int(row['bandwidth'][:3])
                else:
                    bandwidth = int(row['bandwidth'][:3]) * 1000
                
                if bandwidth >= min_bandwidth and bandwidth <= max_bandwidth:
                    child_node = core.create_child(bundle, self.META['Part'])
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

        commit_info = self.util.save(root_node, self.commit_hash, 'test', 'Python plugin populated the Velocloud vendor bundles with parts')
        logger.info('committed :{0}'.format(commit_info))
