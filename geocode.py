from geolocation.main import GoogleMaps

from pyexcel_xlsx import get_data
from pyexcel_xlsx import save_data
from collections import OrderedDict
import json
data = get_data("bikeinfo.xlsx")
# print data['Blad1']
n=0
sheet={"Sheet 1":[["x","y","postcode","city","formatted_address"]]}
google_maps = GoogleMaps(api_key='AIzaSyAMQRmEbUVvs7jZ5tW8hujZzUqqFy4syXc')
for each in data['Blad1']:
    n=n+1
    if n>1:
        print each[1]
        print each[0]
        try:
            location = google_maps.search(lat=each[1], lng=each[0]) # sends search to Google Maps.
            my_location = location.first() # returns only first location.
            print(my_location.city)
            print(my_location.postal_code)
            print(my_location.formatted_address)
            sheet["Sheet 1"].append([each[1], each[0],my_location.postal_code,my_location.city,my_location.formatted_address])
        except:
            sheet["Sheet 1"].append([each[1], each[0],"error","error","error"])
save_sheet = OrderedDict()
save_sheet.update(sheet)
save_data("out_put1.xlsx", save_sheet)
