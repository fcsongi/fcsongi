# from os import access
# import requests, json

# base_url = "https://uat-api.pccwg-osp.com"

# def authenticate():
#     path = "/token"
#     payload = {
#         "grant_type":"client_credentials",
#         "scope":"openid"
#     }
#     headers = {
#         'Authorization': 'Basic SmY0RUs0WWRKb21jNmY0amNKb1lkdHA3Qk80YTpxWVZnTEIybFhodjdiNVZXZmZZZ0d1Y2NTWElh'
#     }

#     response = requests.post(base_url + path, headers=headers, data=payload)
#     return response.json()["access_token"]



# def check_address(data,access_token):
#     path = "/mef/geographicAddressManagement/2.0.0/geographicAddressValidation"
#     payload = {
#         "formattedAddress":{
#             "addrLine1":data["address"],
#             "city":data["city"],
#             "postcode":data["zipcode"],
#             "country":data["country"],
#         }
#     }

#     headers = {
#         "Authorization": "Bearer " + access_token
#     }

#     response = requests.post("https://www.pccwg-osp.com/mef/geographicAddressManagement/2.0.0/geographicAddressValidation", headers=headers, json=payload)
#     print(response)
#     print(response.json())
    
# def create_pccw_quote(data):
#     access_token = authenticate()
#     check_address(data=data,access_token=access_token)
    
# data = {
#     "country":"United States",
#     "zipcode":"46143",
#     "city":"Greenwood",
#     "address":"Endress Place 2350"
#     }

# create_pccw_quote(data)