from bs4 import BeautifulSoup
import requests
import re
import datetime
import pandas as pd

when2meet_link = input()
when2meet_page = requests.get(when2meet_link)
soup = BeautifulSoup(when2meet_page.content, "html.parser")

def get_time(soup):
    table_data = soup.find(text=re.compile('TimeOfSlot*'))
    unix_time = re.compile(r"TimeOfSlot\[\d+\]=(\d+);").findall(table_data)
    normal_time = []
    for unix in unix_time:
        time_stamp = datetime.datetime.fromtimestamp(int(unix))
        time = time_stamp.strftime('%d %b %H:%M')
        normal_time.append(time)
    return normal_time

def get_nameid_names(soup):
    availability_grids = soup.find('div', {'id': 'AvailabilityGrids'})
    script_tags = availability_grids.find_all('script', {'type': 'text/javascript'})
    ids = re.findall(r'PeopleIDs\[\d+\] = (\d+);', script_tags[0].contents[0])
    names = re.findall(r"PeopleNames\[\d+\] = '([ a-zA-Z]{2,})';", script_tags[0].contents[0])
    nameid_names = dict(zip(ids,names))
    return nameid_names

def get_slot_name(soup, nameid_names):
    table_data = soup.find(text=re.compile('TimeOfSlot*'))
    dataset = re.compile(r"AvailableAtSlot\[(\d+)\]\.push\((\d+)\)").findall(table_data)
    slot_names = [(slot, nameid_names[nameid]) for slot, nameid in dataset] #Convert ID to Names
    return slot_names

def get_dataset(index, columns, values):
    export_dataset = pd.DataFrame(index = index, columns = columns.values())
    for slot,names in values:
        col_pos = export_dataset.columns.get_loc(names)
        export_dataset.iloc[int(slot),col_pos]=1
    export_dataset = export_dataset.fillna(0)
    return export_dataset 

def multiply_values_in_every_nrows(df, total_rows):
    for i in range(0, len(df)+total_rows, total_rows):
        df.iloc[i:i+total_rows] = df.iloc[i:i+total_rows].apply(lambda x: x.prod())
    return df

normal_time = get_time(soup)
nameid_names = get_nameid_names(soup)
slot_names = get_slot_name(soup, nameid_names)
export_dataset = get_dataset(normal_time, nameid_names, slot_names)

export_dataset = multiply_values_in_every_nrows(export_dataset, 4)
export_dataset = export_dataset.iloc[3::4]
export_dataset.index = export_dataset.index.str.replace('45','00')
export_dataset.to_excel("Schedule.xlsx", engine = 'openpyxl')