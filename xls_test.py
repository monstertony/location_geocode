from pyexcel_xlsx import get_data
import json
data = get_data("bikeinfo.xlsx")
print data
# print data['Blad1']
for each in data['sheet']:
    print each