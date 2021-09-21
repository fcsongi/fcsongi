"""
This is where the implementation of the plugin code goes.
The Create_IOR_Estimator-class is imported from both run_plugin.py and run_debug.py
"""

import pandas
import json
import sys
import logging
from webgme_bindings import PluginBase

# Setup a logger
logger = logging.getLogger('Create_IOR_Estimator')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

class Create_IOR_Estimator(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node

        name = core.get_attribute(active_node, 'name')

        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))

        #loads excel as json
        excel_df = pandas.read_excel('./imports/IORestimator.xlsx', sheet_name='IOR Delivery estimator',dtype=object)
        result = excel_df.to_json(orient="records")
        data = json.loads(result)

        IOR_meta = self.META["Estimation"]

        i = 1
        for row in data:
            estimator = core.create_child(active_node,IOR_meta)
            position_item = core.get_registry(active_node,'position')
            position_item['y'] += 50 * i
            core.set_registry(estimator,"position",position_item)
            core.set_attribute(estimator,"name",row["Country code"])
            core.set_attribute(estimator,"Lane fee",row["Lane fee"])
            core.set_attribute(estimator,"Shipping cost multiplier",row["Shipping cost multiplier"])
            core.set_attribute(estimator,"Tax and duties multiplier",row["Tax and duties multiplier"])
            i+=1

        commit_info = self.util.save(root_node, self.commit_hash, 'master', 'Python plugin updated the model')
        logger.info('committed :{0}'.format(commit_info))
