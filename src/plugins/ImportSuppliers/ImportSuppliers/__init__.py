"""
This is where the implementation of the plugin code goes.
The ImportSuppliers-class is imported from both run_plugin.py and run_debug.py
"""
import sys
import json
import logging
from webgme_bindings import PluginBase

# Setup a logger
logger = logging.getLogger('ImportSuppliers')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

def create_connections(self, supplier_name, supplier_node, product_node, connection_type, country_codes):
    core = self.core
    country_nodes = load_all_countries(self)

    for country_code in country_codes:
        for country_node in country_nodes:
            if country_code == core.get_attribute(country_node, "isoCode3"):
                connection_node = core.create_child(product_node, self.META[connection_type])
                core.set_attribute(connection_node, "name", supplier_name + "-" + core.get_attribute(country_node, "name"))
                core.set_pointer(connection_node, "src", supplier_node)
                core.set_pointer(connection_node, "dst", country_node)


def create_products(self, supplier_node, supplier):
    core = self.core
    j=0
    for product in supplier["products"]:
        product_node = core.create_child(supplier_node, self.META[product["product_name"]])
        core.set_registry(product_node, "position", {'x':200,'y':75 + 50 * j})
        core.set_attribute(product_node, "builder", json.dumps(product["builder"], indent=4))
        #Removing the last char of the key because the product is a container and a connection is one entity
        if product["product_name"] == "Connectivities":
            connection_type = "Connectivity"
        else:
            connection_type = product["product_name"][:-1]

        create_connections(self, supplier_name=supplier["supplier_name"], supplier_node=supplier_node, product_node= product_node,connection_type = connection_type, country_codes=product["supported_countries_codes"])
        j+=1

def load_all_countries(self):
    core = self.core
    regionNodes = core.load_children(self.META["Countries"])
    countryNodes = []
    if regionNodes:
        for regionNode in regionNodes:
            if core.get_base_type(regionNode) == self.META["Region"]:
                countryNodes += core.load_children(regionNode)
        return countryNodes
    else:
        print("There are no regions in the database")

class ImportSuppliers(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node
        name = core.get_attribute(active_node, 'name')
        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))
        
        with open("./imports/suppliers.json") as f:     #loading regions json file into data list
            suppliers = json.load(f)

        i=0
        for supplier in suppliers:
            supplier_node = core.create_child(active_node, self.META["Supplier"])
            core.set_attribute(supplier_node, "name", supplier["supplier_name"])
            core.set_attribute(supplier_node, "baseURL", supplier["baseURL"])
            core.set_attribute(supplier_node, "builder", json.dumps(supplier["builder"],indent=4))
            core.set_registry(supplier_node, "position", {'x':200,'y':75 + 50 * i})

            create_products(self,supplier_node=supplier_node, supplier=supplier)
            i+=1

        commit_info = self.util.save(root_node, self.commit_hash, 'test', 'Python plugin updated the model')
        logger.info('committed :{0}'.format(commit_info))
