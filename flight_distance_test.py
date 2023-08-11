from math import asin, cos, radians, sin, sqrt
import os
import pandas as pd
#test
os.system('cls')

f = pd.ExcelFile('flight_distance_test.xlsx')

r = 3956 # Radius of Earth (miles)
distance = []

for i, row in f.parse(sheet_name = 'Flight Distance Test').iterrows():
    lat1_rad = radians(row['Departure_lat'])
    lon1_rad = radians(row['Departure_lon'])
    lat2_rad = radians(row['Arrival_lat'])
    lon2_rad = radians(row['Arrival_lon'])

    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad
    
    a = sin(dlat / 2)**2 + cos(lat1_rad) * cos(lat2_rad) * sin(dlon / 2)**2
    c = 2 * asin(sqrt(a))

    d = c * r

    distance.append(round(d, 2))

data = {'Distance (mi)': distance}

df = pd.concat([pd.read_excel('flight_distance_test.xlsx'), \
    pd.DataFrame(data, columns = ['Distance (mi)'])], axis = 1)

df.to_excel('output.xlsx', index = False)
os.startfile('output.xlsx')