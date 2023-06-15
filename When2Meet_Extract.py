from bs4 import BeautifulSoup
import requests
import re
import datetime
import pandas as pd

when2meet_link = input()
when2meet_page = requests.get(when2meet_link)

class GetData:
    def __init__(self, soup):
        self.soup = soup
        self.normal_time = self.get_time()
        self.nameid_names = self.get_nameid_names()

    def get_time(self):
        # When2Meet has its time in UNIX format in the following structure: TimeOfSlot[##]=UNIX Time;.
        table_data = self.soup.find(text=re.compile('TimeOfSlot*'))
        unix_time = re.compile(r"TimeOfSlot\[\d+\]=(\d+);").findall(table_data)
        normal_time = []
        for unix in unix_time:
            time_stamp = datetime.datetime.fromtimestamp(int(unix))
            time = time_stamp.strftime('%d %b %H:%M')
            normal_time.append(time)
        return normal_time

    def get_nameid_names(self):
        # In <div id=AvailabilityGrids><script type="text/javascript">, When2Meet lists the unique IDs and the corresponding names.
        availability_grids = self.soup.find('div', {'id': 'AvailabilityGrids'})
        script_tags = availability_grids.find_all('script', {'type': 'text/javascript'})
        ids = re.findall(r'PeopleIDs\[\d+\] = (\d+);', script_tags[0].contents[0])
        names = re.findall(r"PeopleNames\[\d+\] = '([ a-zA-Z]{2,})';", script_tags[0].contents[0])
        nameid_names = dict(zip(ids,names))
        return nameid_names

    def get_slot_name(self):
        # When2Meet uses the following structure to label who is available at a time slot: AvailableAtSlot[Index].push(Unique ID);
        table_data = self.soup.find(text=re.compile('TimeOfSlot*'))
        dataset = re.compile(r"AvailableAtSlot\[(\d+)\]\.push\((\d+)\)").findall(table_data)
        slot_names = [(slot, self.nameid_names[nameid]) for slot, nameid in dataset] #Convert ID to Names
        return slot_names

def get_dataset(index_name, column_name, availability):
    export_dataset = pd.DataFrame(index = index_name, columns = column_name.values())
    for slot,names in availability:
        col_pos = export_dataset.columns.get_loc(names)
        export_dataset.iloc[int(slot),col_pos]=1
    export_dataset = export_dataset.fillna(0)
    return export_dataset 

def multiply_values_in_every_nrows(df, number_of_rows):
    for index in range(0, len(df)+number_of_rows, number_of_rows):
        df.iloc[index:index+number_of_rows] = df.iloc[index:index+number_of_rows].apply(lambda row: row.prod())
    return df

soup = GetData(BeautifulSoup(when2meet_page.content, "html.parser"))
slot_names = soup.get_slot_name()
export_dataset = get_dataset(soup.normal_time, soup.nameid_names, slot_names)

EVERY_HOUR = 4 #When2Meet blocks are in 15 minute intervals.
export_dataset = multiply_values_in_every_nrows(export_dataset, EVERY_HOUR) #Checks who is available for the full hour.
export_dataset = export_dataset.iloc[::EVERY_HOUR]
export_dataset.to_excel("Schedule.xlsx", engine = 'openpyxl')