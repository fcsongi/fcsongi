import json, requests

def search_supplier_for_services(self,searched_product,site):
    core = self.core
    suppliers = core.load_children(self.META["Suppliers"])
    for supplier in suppliers:
        products = core.load_children(supplier)
        for product in products:
            if searched_product == core.get_attribute(product,"name"):
                connections = core.load_children(product)
                for connection in connections:
                    if site["Country"] == core.get_attribute(connection, "name").split("-")[1]:
                        #send api call
                        pass
    return site



def auth(self, supplierNode):
    core = self.core
    authBuilder = core.get_attribute(supplierNode, "auth")
    baseURL = core.get_attribute(supplierNode, "baseURL")
    response = requests.request(authBuilder["authMethod"], url=baseURL + authBuilder["authPath"], json = authBuilder["payload"], headers=authBuilder["headers"])
    if response.status_code == requests.codes.ok:
        return response.json()
    else:
        with open("./errors/authError/" + "projectName" + ".txt", "a") as f:
            f.write(core.get_attribute(supplierNode, "name") + "\n")
            f.write(json.dumps(response.json(),indent = 4))
            f.close()
            return 


def callSupplierApi(self, supplierNode, productNode, site):
    core = self.core
    authResponse = auth(self,supplierNode)
    
    

    pass