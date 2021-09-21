"""
This is where the implementation of the plugin code goes.
The SupplierLogic-class is imported from both run_plugin.py and run_debug.py
"""
import json, requests, time,  os
from datetime import datetime
import base64
import logging, sys
from webgme_bindings import PluginBase

# # Setup a logger
logger = logging.getLogger('SupplierLogic')
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stdout)  # By default it logs to stderr..
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

site = {}
authentication = {}
response = {}

#This function returns the active_node's (a connection's) targeted node by the given terget_pointer_name ('src'/'dst')
def get_conn_obj_meta(core, active_node, target_pointer_name, connected_type, mixin = False):
    """
    This function searches the connection node which has the specified connection type and connects with the active node
    """
    if core.is_connection(active_node):
        for pointer_name in core.get_pointer_names(active_node):
            if pointer_name == target_pointer_name:
                relative_path=core.load_pointer(active_node, pointer_name)['nodePath']
                root_node = core.get_root(active_node)
                connected_node =core.load_by_path(root_node, relative_path)
                if not(mixin):
                    if connected_type == None or  core.is_instance_of(connected_node, connected_type):
                        return connected_node
                else:
                    if connected_type == None or  core.is_type_of(connected_node, connected_type):
                        return connected_node

#This function returns a list with the nodes (peer_type) that are connected with the given node
def get_peer_objects_meta(core, active_node, pointer_target, connector_type, peer_type, mixin = False, return_connections = False):
    """
    This function searches the node which has the specified connection type and connects with the active node
    """
    ret_val=[]
    peer_node_paths = core.get_collection_paths(active_node,pointer_target)
    for peer_node_path in peer_node_paths:
        root_node = core.get_root(active_node)
        peer_con_node =core.load_by_path(root_node, peer_node_path)
        if connector_type == None or core.is_instance_of(peer_con_node, connector_type):
            if pointer_target == 'dst':
                theother_pointer_target ='src'
            else:
                theother_pointer_target ='dst'
            peer_obj = get_conn_obj_meta(core, peer_con_node, theother_pointer_target, peer_type, mixin)
            if peer_obj != None:
                if (return_connections == False):
                    ret_val.append(peer_obj)
                else:
                    ret_val.append([peer_con_node, peer_obj])
    return ret_val

def getProductNodeOfSupplier(self, supplierNode, product):
    """
    This function searches the given product type in the supplier node
    Returns the found node if finds it
            None if it was not found
    """
    logger.info("Searching for specified product in supplier")
    core = self.core
    if product == "Connectivity":
        product = "Connectivities"
    else:
        product = product + "s"
    for productNode in core.load_children(supplierNode):
        if core.get_attribute(productNode, "name") == product:
            logger.info("Product found")
            return productNode
    
    logger.info("Could not find " + product + " in the given node.")
    logError(message="Could not find " + product + " in in the given node.", supplierName=core.get_attribute(supplierNode,"name"))

def logError(message, supplierName= "", requestType= "", url= "", data= {}, headers= {} ):
    """
    This function logs the error to the coresponding filesystem 
    """
    logger.info("An error occured. Message: " + str(message))
    os.system("mkdir -p ./errors/api/" + str(site["quoteId"]))
    with open("./errors/api/" + site["quoteId"] + "/" + site["siteId"] + ".txt", "a") as f:
        f.write("An error occured at " + str(datetime.now()) + " in " + site["project_name"] + "\n")
        f.write("Supplier name: " + supplierName + "\n")
        f.write("Site information:\n")
        f.write(json.dumps(site, indent=4))
        f.write("\n")
        f.write("Authentication:\n")
        f.write(json.dumps(authentication, indent=4))
        f.write("\n")
        f.write("Message: \n")
        f.write(str(message) + "\n")
        f.write("Response: \n")
        f.write(response.text)
        f.write("Api request\n")
        f.write("url: " + url + "\n")
        f.write("Request type: " + requestType + "\n")
        f.write("Headers: \n")
        f.write(json.dumps(headers, indent=4))
        f.write("\n")
        f.write("Data: \n")
        f.write(json.dumps(data, indent=4))
        f.write("\n")



def setKeyOfObject(core, process_dictionary):
    """
    This function sets the key of a dictionary to the given value
    Usage: process dictionary = {
            "dictionary you want to modify":{
                    "key1":"value1"
                    "key2":"value2"
                }
            }
    Possible dictionaries: authentication, site
    Returns None
    """
    global authentication, site
    processJsonCommand(core, process_dictionary)
    obj = list(process_dictionary.keys())[1]
    if obj == "authentication":
        for key, value in process_dictionary[obj].items():
            authentication[key] = value
    elif obj == "site":
        for key, value in process_dictionary[obj].items():
            site[key] = value
    else:
        logger.info("The given process_response variable is unknown. " + obj + "\n")
        logError(message="The given process_response variable is unknown. " + obj + "\n")

def processJsonCommand(core, obj):
    """
    This command processes the given json from webgme builder
    Substitues the given functions and special variables in the string
    Example: "authentication":{
        "token":"getKeyFrom:['acess_token'],response.json()",       -> will be the value of response.json()['access_token']
        "Authorization":"eval:authentication['tokenType'] + ' ' + authentication['token']"     -> will evaluate the given string
    }
    """
    if type(obj) == dict:
        for key in obj:
            evaluated_obj = processJsonCommand(core, obj[key])
            if evaluated_obj is not None:
                obj[key] = evaluated_obj

    elif type(obj) == list:
        for i in range(len(obj)):
            evaluated_obj = processJsonCommand(core, obj[i])
            if evaluated_obj is not None:
                obj[i] = evaluated_obj
    
    elif type(obj) == str:
        if obj.startswith("eval:"):
            obj = obj.replace("eval:","")
            return eval(obj)
        elif obj.startswith("getKeyFrom:"):
            #preparing request
            obj = obj.replace("getKeyFrom:", "")
            obj = obj.split(",")
            returnObj = None
            try:
                returnObj = eval(obj[1] + obj[0])
            except KeyError:
                if obj[0] == "['portType']":
                    if int(site["bandwidth"]) <= 10:
                        portType = "10M Ethernet"
                    elif int(site["bandwidth"]) <= 100:
                        portType = "100M Ethernet"
                    elif int(site["bandwidth"]) <= 1000:
                        portType = "1G Ethernet"
                    elif int(site["bandwidth"]) <= 10000:
                        portType = "10G Ethernet"
                    elif int(site["bandwidth"]) <= 100000:
                        portType = "100G Ethernet"
                    returnObj = portType
            finally:
                if returnObj is not None:
                    return returnObj
                else:
                    logger.info("Could not get key: " + obj[0] + " from " + obj[1] + " \n")
                    logError(message="Could not get key: " + obj[0] + " from " + obj[1] + " \n")


def getCountryNameFromSupplier():
    """
    Searches the countries name by the site's countryCode2
    Returns the countrie's name
    """
    for country in response.json()["ShipTOCountriesList"]:
        if country["TwoDigitISO"] == site["countryCode2"]:
            return country["Name"]


def auth(self, supplierNode):
    """
    This function retrieves the necessary tokens for the further api communications
    Returns True if authentication succeded
            False if authentication failed
    """
    logger.info("Authenticating...")
    core = self.core
    builder = json.loads(core.get_attribute(supplierNode, "builder"))
    baseURL = core.get_attribute(supplierNode, "baseURL")
    for process in builder:
        headers = process["headers"]
        payload = process["payload"]
        processJsonCommand(core=core,obj=headers)
        processJsonCommand(core=core,obj=payload)
        process_response = process["process_response"]
        send_request(requestType=process["method"], url= baseURL + process["path"], data= payload, payloadType=process["payloadType"],headers= headers, process_response= process_response, supplierName=core.get_attribute(supplierNode, "name"))
        try:
            #response.json() might not be json and causes Exception
            setKeyOfObject(core, process_response[str(response.status_code)])
        except Exception as e:
            logger.info("Authentication failed with " + core.get_attribute(supplierNode, "name"))
            logError(message=e,supplierName=core.get_attribute(supplierNode,"name"))
            return False
    logger.info("Authentication is succesful.")
    logger.info("Authentication dict:")
    logger.info(json.dumps(authentication, indent=4))
    return True
        
def callSupplierApi(self, supplierNode, service):
    """
    This function calls the given supplier's api to retrieve data
    It will modify the site dictionary
    Returns None
    """
    core = self.core
    logger.info("Api communication initiated.")
    if auth(self=self, supplierNode=supplierNode):
        productNode = getProductNodeOfSupplier(self=self, supplierNode=supplierNode, product=service)
        if productNode is not None:
            logger.info("Creating quote")
            builder = json.loads(core.get_attribute(productNode, "builder"))
            baseURL = core.get_attribute(supplierNode, "baseURL")
            for process in builder:
                process_response = process["process_response"]
                headers = process["headers"]
                payload = process["payload"]
                payloadType = process["payloadType"]
                processJsonCommand(core=core,obj=headers)
                processJsonCommand(core=core,obj=payload)
                if send_request(requestType=process["method"], url= baseURL + process["path"], data= payload, payloadType=payloadType, headers= headers, process_response= process_response, supplierName=core.get_attribute(supplierNode, "name")):
                    try:
                        # response.json() might not be json and causes Exception
                        setKeyOfObject(core, process_response[str(response.status_code)])
                    except Exception as e:
                        logger.info("Quote retrieving failed with " + core.get_attribute(supplierNode, "name"))
                        logError(message=e,supplierName=core.get_attribute(supplierNode,"name"))
                        return

                

def searchSuppliersWithServiceInCountry(self, countryNode, service):
    """
    This function searches suppliers that can provide the selected service in the given country node
    Returns a list
    """
    core = self.core
    suppliers = get_peer_objects_meta(core= core, active_node= countryNode, pointer_target="dst", connector_type=self.META[service], peer_type=self.META["Supplier"])
    return suppliers

    
def send_request(requestType, url, data, headers, supplierName, payloadType = None, process_response = {}):
    """
    This function sends the prepared request
    Returns True if succeded
            False if failed
    """
    global response
    j = 0
    for i in range(151):
        try:
            logger.info("Sending request to " + url)
            if payloadType == "json":
                response = requests.request(method=requestType, url= url, json=data, headers=headers)
            elif payloadType == "params":
                response = requests.request(method=requestType, url= url, params=data, headers=headers)
            else:
                response = requests.request(method=requestType, url= url, data=data, headers=headers)
            time.sleep(2)
            logger.info("Response status code: " + str(response.status_code))
            if str(response.status_code) in process_response.keys():
                logger.info("Expected response recieved. " + str(response.status_code))
                logger.info(json.dumps(response.json(), indent=4))
                wrongKeyFlag = False
                if process_response[str(response.status_code)]["checkKeysAndValues"] != {}:
                    for key, value in process_response[str(response.status_code)]["checkKeysAndValues"].items():
                        try:
                            if eval("response.json()" + key) != value:
                                logger.info("Key found but value is not what we expected. Key: " + key + "Value: " + value + " Retrying..")
                                wrongKeyFlag = True
                        except KeyError:
                            logger.info("Key not found in response. Key: " + key + " Response keys: + " + str(response.json().keys()) + " Retrying..")
                            wrongKeyFlag = True
                if wrongKeyFlag is False :
                    return True
            else:
                logError(message="Unhandled status code recieved. Retrying..\nStatus code: " + str(response.status_code) + "\nResponse text: " + response.text + "\n",supplierName=supplierName, url=url,requestType=requestType,data=data,headers=headers)
        except Exception as e:
            delaySec = 30
            logError(message="Exception occured when sending request. Retrying in " + str(delaySec) +" sec..\nError: " + str(e) + "\n", supplierName=supplierName,url=url,requestType=requestType,data=data,headers=headers)
            if j >= 5:
                logError(message="All 3 retries faileed. Exiting loop.\nError: " + str(e) + "\n",supplierName=supplierName,url=url,requestType=requestType,data=data,headers=headers)
                return False
            logger.info("Could not send request to " + supplierName)
            logger.info("Retrying in " + str(delaySec) + " seconds")
            time.sleep(delaySec)
            j+=1
    logger.info("Number of request iteration has been exceeded. Quote recieving failed.")
    logError(message="Number of request iteration has been exceeded. Quote recieving failed.", supplierName=supplierName)
    return False

def initiateApiCallForSite(self,quotingsite,service,syslogger):
    """
    This function initiates the api workflow to the given site
    Directly modifies the site dictionary
    Returns the modified site
    """
    global site
    global logger
    logger=syslogger
    site = quotingsite
    #countryNode path is an example
    countryNode = site["countryNode"]
    if service == "Connectivity":
        site["Supplier OTC cost"] = 0
        site["Supplier MRC cost"] = 0

    logger.info("LOGGER SEARCHIN FOR SUPPLIER NODES")
    print("SEARCHIN FOR SUPPLIER NODES")
    supplierNodes = searchSuppliersWithServiceInCountry(self=self, countryNode=countryNode, service=service)
    logger.info("Found suppliers for the requested country.")
    for supplierNode in supplierNodes:
        callSupplierApi(self=self, supplierNode=supplierNode, service=service)
    print(json.dumps(site, indent=4))
    return site

class SupplierLogic(PluginBase):
    def main(self):
        core = self.core
        root_node = self.root_node
        active_node = self.active_node

        name = core.get_attribute(active_node, 'name')

        logger.info('ActiveNode at "{0}" has name {1}'.format(core.get_path(active_node), name))
        # service = "Connectivity"
        # testsite = {
        #     "countryNode":core.load_by_path(root_node,"/y/MI0/s"),
        #     "project_name":"hmm",
        #     "quoteId":"1234",
        #     "siteId":"4",
        #     "countryCode2":"AT",
        #     "countryCode3":"AUT",
        #     "city":"Linz",
        #     "address":"Linz",
        #     "bandwidth": "100",
        #     "contract_term_month": 36,
        #     "contract_term_year": 3
        # }

        service = "IOR provider"
        testsite = {
            "project_name":"hmm",
            "quoteId":"1234",
            "siteId":"4",
            "countryNode":core.load_by_path(root_node,"/y/MI0/s"),
            "countryCode2":"AT",
            "countryCode3":"AUT",

            "shipment_weight":12,
            "number_of_sites":1,
            "shipment_value":1798.8
        }
        

        testsite = initiateApiCallForSite(self, quotingsite=testsite, service=service, syslogger=logger)

        commit_info = self.util.save(root_node, self.commit_hash, 'test', 'Python plugin updated the model')
        logger.info('committed :{0}'.format(commit_info))