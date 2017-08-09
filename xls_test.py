from pyexcel_xlsx import get_data
import json
data = get_data("coordinates.xlsx")
print data['Blad1']
for each in data['Blad1']:
    print each