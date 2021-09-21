import requests
import base64
import json

url = "https://tecex.force.com/ncpo/services/apexrest/"
endpoint = "PortalLogin"
payload = {'username':'adorjani.arpad@combridge.ro', 'password':'yvc8-C9#!cHdGxW', 'domain':'login'}
r = requests.post(url + endpoint, params=payload, headers={})
print(json.dumps(r.json(),indent=4))

# sessionId = r.json()['sessionId']

# headers = {"Authorization": "Bearer " + sessionId}

# body = {
#     "ClientuserID" : "0051v0000096oTjAAI"
# }
# endpoint = "MyDetails"
# r = requests.post(url + endpoint, json=body, headers=headers)
# print(json.dumps(r.json(),indent=4))

# endpoint = "MyDestinationCountrieslist"
# r = requests.get(url + endpoint, headers=headers)
# print(json.dumps(r.json(), indent=4))

# for country in r.json()["ShipTOCountriesList"]:
#     if country["TwoDigitISO"] == "RO":
#         print(country["Name"])


# client = r.json()['ClientUserIds'][0]

# quotes = [
#     {
#         "estimatedChargableweight":12,
#         "ServiceType":"IOR",
#         "Courier_responsibility":"Client",
#         "ShipFrom":"Romania",
#         "ShipTo":"United States",
#         "ShipmentvalueinUSD":1798.8,
#         "PONumber":"TestPONumber",
#         "Type_of_Goods":"New",
#         "Li_ion_Batteries":"No",
#         "Li_ion_BatteryTypes":"",
#         "NumberOfFinaldeliveries":1,
#         "ShipmentOrder_Packages":[],
#         "FinalDelivieries":[],
#         "Parts":[]
#     }
# ]

# accessToken = base64.b64decode(client['AccessToken'])
# accessToken = accessToken.decode('utf-8')

# payload = {
#         "AccountID" : client['ClientAccID'],
#         "ContactID" : client['ClientContactId'],
#         "AccessToken" : accessToken,
#         "Quotes" : quotes
# }

# endpoint = "CreateQuickQuote"
# r = requests.post(url + endpoint, json=payload, headers=headers)
# print(r)
# print(r.json())

# payload = {
#     # "SOID":r.json()["Id"],
#     "SOID":"a0R1v00000QKRLMEA5",
#     "AccessToken" : accessToken
# }
# endpoint = "CEDetails"
# r = requests.post(url + endpoint, headers=headers, json=payload)
# print(json.dumps(r.json(), indent=4))






# print(r.json()["ShipmentOrder"]["Totalincludingestimateddutiesandtax"])

# r = requests.get(url + "MyDestinationCountrieslist", headers = headers)

# print(json.dumps(r.json(), indent=4))

# IORrequest_row = {
#     "Shipment Order" : "",
#     "Ship To Country" : "",
#     "Shipment Value (USD)" : "",
#     "IOR and Import Compliance" : "",
#     "Admin Fee": "",
#     "International Freight Fee (1)" : "",
#     "Liability Cover Fee (2)" : "",
#     "Total - Customs Brokerage Cost" : "",
#     "Total - Clearance Costs" : "",
#     "Total - Handling Costs" : "",
#     "Total - License Cost" : "",
#     "Estimated - Tax and Duty (3)" : "",
#     "Bank Fees" : "",
#     "Cash Disbursement Fee (4)" : "",
#     "Total Invoice Amount" : "",
#     "Chargeable Weight" : "",
#     "Estimate Transit Time" : "",
#     "Estimate Customs Clearance Time" : "",
#     "Shipping Notes" : ""
# }