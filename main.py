import datetime
import os

import xlrd
import xlsxwriter

orbitz = []
costa = []
princess = []
hal = []
carnival = []
carnival_au = []
cunard = []
rss = []
royal = []
celebrity = []
oceania = []
ncl = []
azamara = []
not_null = []
new = []
workbook = xlrd.open_workbook('Missing Prices -- 2017-06-16 -- By Company.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []


def xldate_to_datetime(xldate):
    try:
        tempDate = datetime.datetime(1900, 1, 1)
        deltaDays = datetime.timedelta(days=int(xldate))
        secs = (int((xldate % 1) * 86400) - 60)
        detlaSeconds = datetime.timedelta(seconds=secs)
        TheTime = (tempDate + deltaDays + detlaSeconds)
        return TheTime.strftime("%m/%d/%Y")
    except ValueError:
        old_value = xldate.split('/')
        new_value = (
            datetime.date(int(old_value[2]), int(old_value[0]), int(old_value[1])) - datetime.date(1899, 12, 30)).days
        tempDate = datetime.datetime(1900, 1, 1)
        deltaDays = datetime.timedelta(days=int(new_value))
        secs = (int((new_value % 1) * 86400) - 60)
        detlaSeconds = datetime.timedelta(seconds=secs)
        TheTime = (tempDate + deltaDays + detlaSeconds)
        return TheTime.strftime("%m/%d/%Y")


for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        if col == 4 or col == 1:
            pass
        else:
            elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
current_ship = ''
current_dest = ''
for d in data:
    if d['dest'] != '':
        current_dest = d['dest']
    if current_ship == '':
        try:
            if d['date'] != '':
                tmp_ship = float(d['date'])
            else:
                continue
        except ValueError:
            current_ship = d['date']
    else:
        try:
            if d['date'] != '':
                return_date = float(d['date'])
                start_date = float(d['ship'])
                interior = d["Interior"]
                oceanview = d['Oceanview']
                balcony = d['Balcony']
                suite = d['Suite']
                if interior != '':
                    if isinstance(interior, str):
                        if interior == 'Sold Out' or interior == 'N/A' or interior == 'CALL':
                            interior = 'N/A'
                        else:
                            interior = int(d['Interior'].split(' | ')[0])
                    elif isinstance(interior, float):
                        interior = int(str(interior).split('.')[0])
                if oceanview != '':
                    if isinstance(oceanview, str):
                        if oceanview == 'Sold Out' or oceanview == 'N/A' or oceanview == 'CALL':
                            oceanview = 'N/A'
                        else:
                            oceanview = int(d['Oceanview'].split(' | ')[0])
                    elif isinstance(oceanview, float):
                        oceanview = int(str(oceanview).split('.')[0])
                if balcony != '':
                    if isinstance(balcony, str):
                        if balcony == 'Sold Out' or balcony == 'N/A' or balcony == 'CALL':
                            balcony = 'N/A'
                        else:
                            balcony = int(d['Balcony'].split(' | ')[0])
                    elif isinstance(balcony, float):
                        balcony = int(str(balcony).split('.')[0])
                if suite != '':
                    if isinstance(suite, str):
                        if suite == 'Sold Out' or suite == 'N/A' or suite == 'CALL':
                            suite = 'N/A'
                        else:
                            suite = int(d['Suite'].split(' | ')[0])
                    elif isinstance(suite, float):
                        suite = int(str(suite).split('.')[0])
                orbitz.append(
                    [current_ship, start_date, return_date, interior, oceanview, balcony, suite, current_dest])
            else:
                continue
        except ValueError:
            if isinstance(d['date'], str):
                current_ship = d['date']
                continue
            else:
                return_date = float(d['date'])
                start_date = float(d['ship'])
                interior = d["Interior"]
                oceanview = d['Oceanview']
                balcony = d['Balcony']
                suite = d['Suite']
                if interior != '':
                    if isinstance(interior, str):
                        if interior == 'Sold Out' or interior == 'N/A' or interior == 'CALL':
                            interior = 'N/A'
                        else:
                            if " | " in interior:
                                interior = int(d['Interior'].split(' | ')[0])
                            else:
                                interior = int(d['Interior'].split(',')[0])
                    elif isinstance(interior, float):
                        interior = int(str(interior).split('.')[0])
                if oceanview != '':
                    if isinstance(oceanview, str):
                        if oceanview == 'Sold Out' or oceanview == 'N/A' or oceanview == 'CALL':
                            oceanview = 'N/A'
                        else:
                            if " | " in oceanview:
                                oceanview = int(d['Oceanview'].split(' | ')[0])
                            else:
                                oceanview = int(d['Oceanview'].split(',')[0])
                    elif isinstance(oceanview, float):
                        oceanview = int(str(oceanview).split('.')[0])
                if balcony != '':
                    if isinstance(balcony, str):
                        if balcony == 'Sold Out' or balcony == 'N/A' or balcony == 'CALL':
                            balcony = 'N/A'
                        else:
                            if " | " in balcony:
                                balcony = int(d['Balcony'].split(' | ')[0])
                            else:
                                balcony = int(d['Balcony'].split(',')[0])
                    elif isinstance(balcony, float):
                        balcony = int(str(balcony).split('.')[0])
                if suite != '':
                    if isinstance(suite, str):
                        if suite == 'Sold Out' or suite == 'N/A' or suite == 'CALL':
                            suite = 'N/A'
                        else:
                            if " | " in suite:
                                suite = int(d['Suite'].split(' | ')[0])
                            else:
                                suite = int(d['Suite'].split(',')[0])
                    elif isinstance(suite, float):
                        suite = int(str(suite).split('.')[0])
                orbitz.append(
                    [current_ship, start_date, return_date, interior, oceanview, balcony, suite, current_dest])
towrite = []
workbook = xlrd.open_workbook('2017-6-17- Costa Cruises.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['InteriorBucketPrice'] != '' and d['SuiteBucketPrice'] != '':
        new.append([d['SailDate'], d['ReturnDate']])
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        costa.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Princess Cruises.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName'], d['ItineraryID'].strip()]
        princess.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Holland America.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        hal.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Carnival US.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        carnival.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Carnival Australia.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        carnival_au.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Cunard Cruises.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['InteriorBucketPrice'] != '' and d['SuiteBucketPrice'] != '':
        new.append([d['SailDate'], d['ReturnDate']])
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        cunard.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- RSSC.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['InteriorBucketPrice'] != '' and d['SuiteBucketPrice'] != '':
        new.append([d['SailDate'], d['ReturnDate']])
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        rss.append(tmp)
workbook = xlrd.open_workbook('2017-6-17 Non - Cruise only price Oceania Cruises.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['InteriorBucketPrice'] != '' and d['SuiteBucketPrice'] != '':
        new.append([d['SailDate'], d['ReturnDate']])
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        oceania.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Royal Caribbean.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['InteriorBucketPrice'] != '' and d['SuiteBucketPrice'] != '':
        new.append([d['SailDate'], d['ReturnDate']])
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        royal.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Norwegian Cruise Line.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['DestinationName'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        ncl.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Celebrity Cruises.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        celebrity.append(tmp)
workbook = xlrd.open_workbook('2017-6-17- Azamara Club Cruises.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)
for d in data:
    if d['DestinationCode'] != '':
        inside = d['InteriorBucketPrice']
        oceanview = d['OceanViewBucketPrice']
        balcony = d['BalconyBucketPrice']
        suite = d['SuiteBucketPrice']
        if isinstance(inside, float):
            inside = int(str(inside).split('.')[0])
        if isinstance(oceanview, float):
            oceanview = int(str(oceanview).split('.')[0])
        if isinstance(balcony, float):
            balcony = int(str(balcony).split('.')[0])
        if isinstance(suite, float):
            suite = int(str(suite).split('.')[0])
        tmp = [d['VesselName'], d['SailDate'], d['ReturnDate'], inside,
               oceanview, balcony, suite, d['DestinationCode'], d['DestinationName']]
        azamara.append(tmp)


def calculate_days(sail_date_param, number_of_nights_param):
    try:
        date = datetime.datetime.strptime(sail_date_param, "%m/%d/%Y")
        calculated = date + datetime.timedelta(days=int(number_of_nights_param))
    except ValueError:
        split = sail_date_param.split('/')
        sail_date_param = split[1] + '/' + split[0] + '/' + split[2]
        date = datetime.datetime.strptime(sail_date_param, "%m/%d/%Y")
        calculated = date + datetime.timedelta(days=int(number_of_nights_param))

    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def convert_to_number(param):
    old_value = param.split('/')
    new_value = (
        datetime.date(int(old_value[2]), int(old_value[0]), int(old_value[1])) - datetime.date(1899, 12, 30)).days
    return new_value


for o in orbitz:
    found = False
    if "Costa " in o[0]:
        sailing = []
        colors = []
        for c in costa:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['Costa Cruises', dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Costa Cruises', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif " Princess" in o[0]:
        sailing = []
        colors = []
        for c in princess:
            if c[9] == '' or c[9] == ' ' or c[9] == '0':
                pass
            else:
                if c[0] == o[0] and o[1] == c[1]:
                    original = xldate_to_datetime(o[2])
                    downloaded = xldate_to_datetime(c[2])
                    if c[9] == "-1":
                        original = calculate_days(original, '-1')
                    elif c[9] == '1':
                        original = calculate_days(original, '1')
                    o[2] = original
                    c[2] = downloaded
                    o[2] = convert_to_number(o[2])
                    c[2] = convert_to_number(c[2])
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['Princess Cruises', dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Princess Cruises', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif "Amsterdam" in o[0] or 'Eurodam' in o[0] or 'Koningsdam' in o[0] or 'Maasdam' in o[0] or 'Nieuw Amsterdam' in \
            o[0] or 'Noordam' in o[0] or 'Oosterdam' in o[0] or 'Prinsendam' in o[0] or 'Rotterdam' in o[
        0] or 'Veendam' in o[0] or 'Volendam' in o[0] or 'Westerdam' in o[0] or 'Zaandam' in o[0] or 'Zuiderdam' in o[
        0]:
        sailing = []
        colors = []
        for c in hal:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['Holland America', dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Holland America', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif "Carnival Legend" in o[0] or "Carnival Spirit" in o[0]:
        sailing = []
        colors = []
        for c in carnival_au:
            if o[0].split()[1] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ["Carnival", dn, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                if 'Elation' in o[0]:
                    print(o)
                    print(c)
                    print(colors)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Carnival', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif "Carnival " in o[0]:
        sailing = []
        colors = []
        for c in carnival:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['Carnival US', dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Carnival US', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)

    elif "Queen Elizabeth" in o[0] or 'Queen Mary 2' in o[0] or 'Queen Victoria' in o[0]:
        sailing = []
        colors = []
        for c in cunard:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['Cunard', dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                print(o)
                print(c)
                print(colors)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Cunard', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif "Celebrity " in o[0]:
        sailing = []
        colors = []
        for c in celebrity:
            if o[0].split()[1] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['Celebrity', dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Celebrity', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif "Norwegian" in o[0] or 'Pride of America' in o[0]:
        sailing = []
        colors = []
        for c in ncl:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['Norwegian', dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Norwegian', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif o[0] == 'Regatta' or o[0] == 'Insignia' or o[0] == 'Sirena' or o[0] == 'Marina' or o[0] == 'Nautica' or o[
        0] == 'Riviera':
        sailing = []
        colors = []
        for c in oceania:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['Oceania', dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Oceania', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif "Azamara " in o[0]:
        sailing = []
        colors = []
        for c in azamara:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ["Azamara", dc, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(o[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['Azamara', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif " of the Seas" in o[0]:
        sailing = []
        colors = []
        for c in royal:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ["Royal", dn, o[0], o[1], o[2]]
                if o[3] == 'N/A' or c[3] == 'N/A':
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append(c[3])
                        colors.append('yellow')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append(o[3])
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append(c[3])
                        colors.append('#23F014')
                    else:
                        temp.append(c[3])
                        colors.append('yellow')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append(c[4])
                        colors.append('yellow')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append(o[4])
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append(o[4])
                        colors.append('#23F014')
                    else:
                        temp.append(c[4])
                        colors.append('yellow')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append(c[5])
                        colors.append('yellow')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append(o[5])
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append(o[5])
                        colors.append('#23F014')
                    else:
                        temp.append(c[5])
                        colors.append('yellow')
                if o[6] == 'N/A' or c[6] == 'N/A':
                    if o[6] == "N/A" and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                    elif o[6] == "N/A" and c[6] != 'N/A':
                        temp.append(c[6])
                        colors.append('yellow')
                    elif o[6] != 'N/A' and c[6] == 'N/A':
                        temp.append(o[6])
                        colors.append('#23F014')
                else:
                    if o[6] == c[6]:
                        temp.append(o[6])
                        colors.append('#23F014')
                    else:
                        temp.append(c[6])
                        colors.append('yellow')
                temp[1] = (o[7])
                sailing.append(temp)
                sailing.append(colors)
                towrite.append(sailing)
                break
            else:
                found = False
        if found:
            pass
        else:
            sailing.append(['Royal', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    elif "Seven Seas " in o[0]:
        sailing = []
        colors = []
        for c in rss:
            if o[0] == c[0] and o[1] == c[1] and o[2] == c[2]:
                found = True
                dc = ''
                dn = ''
                temp = ['RSSC', dc, o[0], o[1], o[2]]
                # Option with more prices
                # if o[3] == 'N/A' or c[3] == 'N/A':
                #     if o[3] == "N/A" and c[3] == 'N/A':
                #         temp.append(o[3])
                #         colors.append('#23F014')
                #     elif o[3] == "N/A" and c[3] != 'N/A':
                #         temp.append(c[3])
                #         colors.append('yellow')
                #     elif o[3] != 'N/A' and c[3] == 'N/A':
                #         temp.append(o[3])
                #         colors.append('#23F014')
                # else:
                #     if o[3] == c[3]:
                #         temp.append(c[3])
                #         colors.append('#23F014')
                #     else:
                #         temp.append(c[3])
                #         colors.append('yellow')
                # if o[4] == 'N/A' or c[4] == 'N/A':
                #     if o[4] == "N/A" and c[4] == 'N/A':
                #         temp.append(o[4])
                #         colors.append('#23F014')
                #     elif o[4] == "N/A" and c[4] != 'N/A':
                #         temp.append(c[4])
                #         colors.append('yellow')
                #     elif o[4] != 'N/A' and c[4] == 'N/A':
                #         temp.append(o[4])
                #         colors.append('#23F014')
                # else:
                #     if o[4] == c[4]:
                #         temp.append(o[4])
                #         colors.append('#23F014')
                #     else:
                #         temp.append(c[4])
                #         colors.append('yellow')
                # if o[5] == 'N/A' or c[5] == 'N/A':
                #     if o[5] == "N/A" and c[5] == 'N/A':
                #         temp.append(o[5])
                #         colors.append('#23F014')
                #     elif o[5] == "N/A" and c[5] != 'N/A':
                #         temp.append(c[5])
                #         colors.append('yellow')
                #     elif o[5] != 'N/A' and c[5] == 'N/A':
                #         temp.append(o[5])
                #         colors.append('#23F014')
                # else:
                #     if o[5] == c[5]:
                #         temp.append(o[5])
                #         colors.append('#23F014')
                #     else:
                #         temp.append(c[5])
                #         colors.append('yellow')
                # if o[6] == 'N/A' or c[6] == 'N/A':
                #     if o[6] == "N/A" and c[6] == 'N/A':
                #         temp.append(o[6])
                #         colors.append('#23F014')
                #     elif o[6] == "N/A" and c[6] != 'N/A':
                #         temp.append(c[6])
                #         colors.append('yellow')
                #     elif o[6] != 'N/A' and c[6] == 'N/A':
                #         temp.append(o[6])
                #         colors.append('#23F014')
                # else:
                #     if o[6] == c[6]:
                #         temp.append(o[6])
                #         colors.append('#23F014')
                #     else:
                #         temp.append(c[6])
                #         colors.append('yellow')
                # temp[1] = (c[7])
                # sailing.append(temp)
                # sailing.append(colors)
                # towrite.append(sailing)

                if o[3] == 'N/A' or c[3] == 'N/A':  # option with suite price only
                    if o[3] == "N/A" and c[3] == 'N/A':
                        temp.append('')
                        colors.append('#23F014')
                    elif o[3] == "N/A" and c[3] != 'N/A':
                        temp.append('')
                        colors.append('23F014')
                    elif o[3] != 'N/A' and c[3] == 'N/A':
                        temp.append('')
                        colors.append('#23F014')
                else:
                    if o[3] == c[3]:
                        temp.append('')
                        colors.append('#23F014')
                    else:
                        temp.append('')
                        colors.append('23F014')
                if o[4] == 'N/A' or c[4] == 'N/A':
                    if o[4] == "N/A" and c[4] == 'N/A':
                        temp.append('')
                        colors.append('#23F014')
                    elif o[4] == "N/A" and c[4] != 'N/A':
                        temp.append('')
                        colors.append('23F014')
                    elif o[4] != 'N/A' and c[4] == 'N/A':
                        temp.append('')
                        colors.append('#23F014')
                else:
                    if o[4] == c[4]:
                        temp.append('')
                        colors.append('#23F014')
                    else:
                        temp.append('')
                        colors.append('23F014')
                if o[5] == 'N/A' or c[5] == 'N/A':
                    if o[5] == "N/A" and c[5] == 'N/A':
                        temp.append('')
                        colors.append('#23F014')
                    elif o[5] == "N/A" and c[5] != 'N/A':
                        temp.append('')
                        colors.append('23F014')
                    elif o[5] != 'N/A' and c[5] == 'N/A':
                        temp.append('')
                        colors.append('#23F014')
                else:
                    if o[5] == c[5]:
                        temp.append('')
                        colors.append('#23F014')
                    else:
                        temp.append('')
                        colors.append('23F014')
                if "Navigator" in o[0]:
                    if o[6] == 'N/A' or c[4] == 'N/A':
                        if o[6] == "N/A" and c[4] == 'N/A':
                            temp.append(o[6])
                            colors.append('#23F014')
                        elif o[6] == "N/A" and c[4] != 'N/A':
                            temp.append(c[4])
                            colors.append('yellow')
                        elif o[6] != 'N/A' and c[4] == 'N/A':
                            temp.append(o[6])
                            colors.append('#23F014')
                    else:
                        if o[6] == c[4]:
                            temp.append(o[6])
                            colors.append('#23F014')
                        else:
                            temp.append(c[4])
                            colors.append('yellow')
                    temp[1] = (o[7])
                    sailing.append(temp)
                    sailing.append(colors)
                    towrite.append(sailing)
                else:
                    if o[6] == 'N/A' or c[5] == 'N/A':
                        if o[6] == "N/A" and c[5] == 'N/A':
                            temp.append(o[6])
                            colors.append('#23F014')
                        elif o[6] == "N/A" and c[5] != 'N/A':
                            temp.append(c[6])
                            colors.append('yellow')
                        elif o[6] != 'N/A' and c[5] == 'N/A':
                            temp.append(o[6])
                            colors.append('#23F014')
                    else:
                        if o[6] == c[5]:
                            temp.append(o[6])
                            colors.append('#23F014')
                        else:
                            temp.append(c[5])
                            colors.append('yellow')
                    temp[1] = (o[7])
                    sailing.append(temp)
                    sailing.append(colors)
                    towrite.append(sailing)
                # option with suite price only
                break
            else:
                found = False
        if found:
            pass
        else:

            sailing.append(['RSSC', o[7], o[0], o[1], o[2], o[3], o[4], o[5], o[6]])
            sailing.append(['cyan', 'cyan', 'cyan', 'cyan'])
            towrite.append(sailing)
    else:
        print("somethign else with", o)


def write_ignore(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    # path_to_file = userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(
    #     now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Costa Cruises Ignore List.xlsx'
    # if not os.path.exists(userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(
    #         now.month) + '-' + str(now.day)):
    #     os.makedirs(
    #         userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook("Latest test.xlsx")

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 18)
    worksheet.set_column("B:B", 10)
    worksheet.set_column("C:C", 18)
    worksheet.set_column("D:D", 10)
    worksheet.set_column("E:E", 10)
    worksheet.set_column("F:F", 10)
    worksheet.set_column("G:G", 10)
    worksheet.set_column("H:H", 10)
    worksheet.set_column("I:I", 10)
    worksheet.write('A1', 'Company', bold)
    worksheet.write('B1', 'DestinationCode', bold)
    worksheet.write('C1', 'VesselName', bold)
    worksheet.write('D1', 'SailDate', bold)
    worksheet.write('E1', 'ReturnDate', bold)
    worksheet.write('F1', 'InteriorBucketPrice', bold)
    worksheet.write('G1', 'OceanViewBucketPrice', bold)
    worksheet.write('H1', 'BalconyBucketPrice', bold)
    worksheet.write('I1', 'SuiteBucketPrice', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    date_format = workbook.add_format({'num_format': 'm/d/yyyy'})
    date_format.set_border()
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    date_format.set_align("center")
    money_format.set_border()
    date_format.set_border()

    for ship_entry in data_array:
        column_count = 0
        for en in ship_entry[0]:
            if column_count == 0:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_align("center")
                ordinary_number.set_border()
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 1:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_border()
                ordinary_number.set_align("center")
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 2:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_align("center")
                ordinary_number.set_border()
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 3:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_align("center")
                ordinary_number.set_border()
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                worksheet.write_number(row_count, column_count, en, date_format)
            if column_count == 4:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_align("center")
                ordinary_number.set_border()
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                worksheet.write_number(row_count, column_count, en, date_format)
            if column_count == 5:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_align("center")
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                centered.set_border()
                try:
                    ordinary_number.set_bg_color(ship_entry[1][0])
                    centered.set_bg_color(ship_entry[1][0])
                    ordinary_number.set_border()
                    worksheet.write_number(row_count, column_count, en, ordinary_number)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 6:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_border()
                ordinary_number.set_align("center")
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                centered.set_border()
                try:
                    ordinary_number.set_bg_color(ship_entry[1][1])
                    centered.set_bg_color(ship_entry[1][1])
                    worksheet.write_number(row_count, column_count, en, ordinary_number)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 7:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_border()
                ordinary_number.set_align("center")
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                centered.set_border()
                try:
                    ordinary_number.set_bg_color(ship_entry[1][2])
                    centered.set_bg_color(ship_entry[1][2])
                    worksheet.write_number(row_count, column_count, en, ordinary_number)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 8:
                ordinary_number = workbook.add_format({"num_format": '#,##0'})
                ordinary_number.set_bold(True)
                ordinary_number.set_border()
                ordinary_number.set_align("center")
                centered = workbook.add_format({'bold': True})
                centered.set_bold(True)
                centered.set_align("center")
                centered.set_border()
                try:
                    ordinary_number.set_bg_color(ship_entry[1][3])
                    centered.set_bg_color(ship_entry[1][3])
                    worksheet.write_number(row_count, column_count, en, ordinary_number)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)

            column_count += 1
        row_count += 1
    workbook.close()
    pass


write_ignore(towrite)
