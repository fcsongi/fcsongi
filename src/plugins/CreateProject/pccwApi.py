import requests, time,json

# base_url = "https://uat-api.pccwg-osp.com" # test env
base_url = "https://api.pccwg-osp.com" # prod env


def authenticate():
  path = "/token"
  payload = {
  "grant_type":"client_credentials",
  "scope":"openid"
  }
  
  headers = {
    'Authorization': 'Basic NFZJREcyU1g5QjZYS0Z1Mlg1Qk5fVlZqWGlnYTpNSnJiTHlvTDlsSnBvSE5LRV9tam14WjlteU1h' #prod env

  }

  # headers = {
  #   'Authorization': 'Basic SmY0RUs0WWRKb21jNmY0amNKb1lkdHA3Qk80YTpxWVZnTEIybFhodjdiNVZXZmZZZ0d1Y2NTWElh' #test env
  # }

  response = requests.post(base_url + path, headers=headers, data=payload)
  print(response.status_code)
  print(json.dumps(response.json(), indent=4))
  return response.json()["access_token"]

def create_internet_access_quote(access_token, quoteId, data):
  path = "/quoteManagement/1.1.0/quote"
  payload = {
    "externalId": quoteId,
    "customer": {
      "id": "COM074"
    }
  }

  quoteItem = []
  row = data
  contract_in_years = int(row["contract_term"]) // 12 
  if int(row["bandwidth"]) <= 10:
    portType = "10M Ethernet"
  elif int(row["bandwidth"]) <= 100:
    portType = "100M Ethernet"
  elif int(row["bandwidth"]) <= 1000:
    portType = "1G Ethernet"
  elif int(row["bandwidth"]) <= 10000:
    portType = "10G Ethernet"
  elif int(row["bandwidth"]) <= 100000:
    portType = "100G Ethernet"

  if row["address"] == "":
    row["address"] = row["city"]

  item = {
    "id":row['siteId'],
    "action":"add",
    "productOffering": {"id": "GIA"},
    "product":{
      "productCharacteristic":[
        { "name": "country", "value": row["isoCode3"] },
        { "name": "city", "value": row["city"] },
        { "name": "address", "value": row["address"] },
        { "name": "portType", "value": portType },
        { "name": "bandwidth", "value": row["bandwidth"] },
        { "name": "bandwidthUnit", "value": "Mbps" },
        { "name": "contractLength", "value": contract_in_years },
        { "name": "contractLengthUnit", "value": "year" }
      ]
    }
  }
  quoteItem.append(item)
  payload["quoteItem"] = quoteItem

  headers = {
      'Authorization': "Bearer " + access_token
  }
  i=0
  while 1:
    response = requests.post(base_url+path, headers=headers, json=payload)
    if response.status_code == 201:
      break
    elif i == 3:
      break
    i+=1
    time.sleep(1)

  return response

def create_pccw_quote(quoteId,data):
  access_token = authenticate()
  print(access_token)
  response = create_internet_access_quote(access_token=access_token, quoteId=quoteId, data=data)
  if response.status_code == 201:
    jsonResponse = response.json()
    for i in range(len(jsonResponse["quoteItem"])):
      for price in jsonResponse["quoteItem"][i]["itemQuoteProductOfferingPrice"]:
        if price["itemQuotePriceAlteration"]["priceType"] == "non-recurrent":
          data["OTC"] = price["itemQuotePriceAlteration"]["price"]["override"]
        if price["itemQuotePriceAlteration"]["priceType"] == "recurrent":
          data["MRC"] = price["itemQuotePriceAlteration"]["price"]["override"]
  print(json.dumps(response.json(), indent=4))
  print(data)
  return data


# #Test
# quoteId = "1234"
# data ={
#         "rowId": 2,
#         "siteId": "3 ",
#         "isoCode3": "AUT",
#         "city": "Linz",
#         "address": "Linz",
#         "bandwidth": "100",
#         "contract_term": 36
#     }

# print(create_pccw_quote(quoteId=quoteId, data=data))

                            