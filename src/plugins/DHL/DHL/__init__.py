"""
This is where the implementation of the plugin code goes.
The DHL-class is imported from both run_plugin.py and run_debug.py
"""

import json
import pandas
import sys
import logging
from webgme_bindings import PluginBase

# Setup a logger
logger = logging.getLogger('DHL')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class DHL(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node

        name = core.get_attribute(active_node, 'name')

        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))

        #loads excel as json
        excel_df = pandas.read_excel('./imports/DHL costs.xlsx', sheet_name='Export costs')
        result = excel_df.to_json(orient="records")
        data = json.loads(result)

        dhl_meta = self.META["DHL"]
        costs_meta = self.META["Costs"]

        dhl_node = core.create_child(active_node,dhl_meta)
        position_item = core.get_registry(active_node,'position')           #getting position of the active node

        i=1
        for row in data:
            child = core.create_child(dhl_node,costs_meta)
            position_item['y'] += 50                                         #changing the position variable
            core.set_registry(child, 'position', position_item)                #changing the child's position
            core.set_attribute(child, 'cost', row['Cost USD'])
            core.set_attribute(child, 'name', row['Weight to zone'])
            if i == 8:
                position_item['y'] -= 400
                position_item['x'] += 150
                i=0
            i+=1

        commit_info = self.util.save(root_node, self.commit_hash, 'master', 'Python plugin populated the Delivery method node')
        logger.info('committed :{0}'.format(commit_info))
