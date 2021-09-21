"""
This is where the implementation of the plugin code goes.
The Vendor_Bundle_generator-class is imported from both run_plugin.py and run_debug.py
"""
import pandas
import json
import sys
import logging
from webgme_bindings import PluginBase

# Setup a logger
logger = logging.getLogger('Vendor_Bundle_generator')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class Vendor_Bundle_generator(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node
        #loads excel as json
        excel_df = pandas.read_excel('./imports/bundles.xlsx',dtype=object)
        result = excel_df.to_json(orient="records")
        bundles = json.loads(result)

        #loading vendors json file into data list
        with open("./imports/vendors.json") as f:
            vendors = json.load(f)

        name = core.get_attribute(active_node, 'name')      #get the name of active_node

        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))  #logging the active_node's name

        #getting the meta of target file which we want to create
        vendor_meta = self.META["Vendor"]
        bundle_meta = self.META["Bundle"]
        customGroup_meta = self.META["CustomEquipmentGroup"]

        customEquipmentGroups = {"SmallEquipments":None, "MediumEquipments":None, "LargeEquipments":None}
        i = 1
        for key in customEquipmentGroups.keys():
            customEquipmentGroup_node = core.create_child(active_node, customGroup_meta)
            core.set_attribute(customEquipmentGroup_node,"name",key)
            position_item = core.get_registry(active_node, 'position')
            position_item['y'] = position_item['y'] + 100 * i
            position_item['x'] += 300
            core.set_registry(customEquipmentGroup_node,'position',position_item)
            customEquipmentGroups[key] = customEquipmentGroup_node
            i+=1

        i = 1
        for vendor in vendors:
            vendor_node = core.create_child(active_node, vendor_meta)
            position_item = core.get_registry(active_node,'position')
            position_item['y'] = position_item['y'] + 100 * i
            core.set_registry(vendor_node, 'position', position_item)
            core.set_attribute(vendor_node, 'name', vendor['name'])
            i+=1

            j = 1
            for bundle in bundles:
                if bundle["vendor"] == core.get_attribute(vendor_node, 'name'):
                    bundle_node = core.create_child(vendor_node, bundle_meta)
                    position_item = core.get_registry(active_node,'position')
                    position_item['y'] = position_item['y'] + 50 * j
                    core.set_registry(bundle_node, 'position', position_item)        
                    core.set_attribute(bundle_node, 'name', bundle["device"])
                    core.set_attribute(bundle_node, 'vendorCode', bundle["vendorCode"])
                    core.set_attribute(bundle_node, 'min_bandwidth', bundle["min_bandwidth"])   
                    core.set_attribute(bundle_node, 'max_bandwidth', bundle["max_bandwidth"])   
                    core.set_attribute(bundle_node, 'cost', bundle["cost"])
                    core.set_attribute(bundle_node, 'weight', round(float(bundle["weight"])+0.5))
                    core.set_attribute(bundle_node, 'discount', bundle["discount"])
                    core.set_attribute(bundle_node, 'description', bundle["description"])
                    j+=1

                    if bundle["size"] == "S":
                        core.add_member(customEquipmentGroups["SmallEquipments"],"MemberBundles", bundle_node)
                    elif bundle["size"] == "M":
                        core.add_member(customEquipmentGroups["MediumEquipments"], "MemberBundles", bundle_node)
                    else:
                        core.add_member(customEquipmentGroups["LargeEquipments"], "MemberBundles", bundle_node)
        
        # creating a commit for the update
        commit_info = self.util.save(root_node, self.commit_hash, 'test', 'Python plugin created vendors and bundles')
        logger.info('committed :{0}'.format(commit_info))
