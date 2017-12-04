import xlrd
import re
import os
import pandas as pd
import datetime
from translitua import translit
import numpy as np
import json
import warnings


monthes_dict = {"січень":"01", "лютий":"02", "березень":"03", "квітень":"04", "травень":"05", "червень":"06", "липень":"07", "серпень":"08", "вересень":"09", "жовтень":"10", "листопад":"11", "грудень":"12"}
STATIONS = ["Придніпровська ТЕС", "Старобешівська ТЕС", 'Слов"янська ТЕС', 'Трипільська ТЕС', 'Зміївська ТЕС', "Луганська ТЕС", "Криворізька ТЕС",
            "Запорізька ТЕС", "Бурштинська ТЕС", "Добротвірська ТЕС", "Ладижинська ТЕС", "Вуглегірська ТЕС", "Курахівська ТЕС", "КЕП «Чернігівська ТЕЦ» ТОВ ФІРМА «ТЕХНОВА»", 'ТОВ "ЄВРО-РЕКОНСТРУКЦІЯ" (Дарницька ТЕЦ)',
            "Сумська ТЕЦ", 'ТОВ "Сумитеплоенерго"', '"ВП Миронівська ТЕС ', 'ПАТ ""ДТЕК Донецькобленерго"", у т.ч.:"', 'ТОВ "Краматорськтеплоенерго"', 'ТОВ «ДВ «Нафтогазовидобувна компанія» (ЄСХАР)',
            'ПАТ "Черкаське хімволокно"' , 'ВП "Черкаська ТЕЦ"', 'ДПЗД «Укрінтеренерго» ВФ «Калуська ТЕЦ»']
year_re = re.compile("(20\d{2})")
file_re = re.compile("\.xlsx?$")

region_marker = "обл."

INPUT_DATA_FOLDER = "../coal_input"
OUTPUT_DATA_FOLDER = "../coal_output"



STATIONS_IDS_FILE = os.path.join(INPUT_DATA_FOLDER, "stations_ids.csv")
STATIONS_INFO_FILE = os.path.join(INPUT_DATA_FOLDER, "stations_info.json")
JSON_DATA_FILE = 'data.json'

MONTHES_DICT = {"01":"січня" , "02":"лютого", "03":"березня", "04":"квітня", "05":"травня", "06":"червня", "07":"липня", "08":"серпня", "09":"вересня", "10":"жовтня", "11":"листопада", "12":"грудня"}

CSV_HEADERS = ["station", "date", "reserve", "min", "max", "planned", "plan_percent", 'completance', "coal_type", "delivery", "spending", "days", "stopped", 'filename']
OUTPUT_FILE = os.path.join(INPUT_DATA_FOLDER,  "coal_reserves_stations.csv")


OUTPUT_STATIC_HEADERS = ["station_id", "latitude", "longitude", "owner", "spending_a_jan2016", "spending_p_jan2016", "spending_g_jan2016"]

NOT_SPACE = re.compile("\S+")

STATIONS_FOLDER = "stations_csvs"

COAL_TYPES_DICTIONARY = {"антрацит":"a", "газове":"g", "пісне":"p"}

STATIONS_POSTFIX = ["_a", "_g", "_p"]


def blank_string_to_null(x):
    if x == "":
        return 0
    else:
        return x

def create_date_string(d):
    d = str(d.date())
    splitted = d.split("-")
    return str(int(splitted[2])) + " " + MONTHES_DICT[splitted[1]] + " " + splitted[0] + " року"

def is_bold(cell):
    font_index = wb.xf_list[cell.xf_index].font_index
    return wb.font_list[font_index].bold == 1

def is_italic(cell):
    font_index = wb.xf_list[cell.xf_index].font_index
    return wb.font_list[font_index].italic == 1

def is_blank(cell):
    return NOT_SPACE.search(str(cell.value)) == None

def get_workbook(filename):
    global reserve_id, days_id
    global wb, sheet, year_month, rows, is_enterprise, date
    reserve_id = None
    days_id = None
    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)
    ncols = sheet.ncols
    nrows = sheet.nrows
    date = xlrd.xldate_as_tuple(sheet.cell(0, 0).value,0)
    date = datetime.datetime(*date[0:6]).strftime("%Y-%m-%d")
    for i in range(2,nrows):
        row = [sheet.cell(i, c_number) for c_number in range(ncols)]
        parse_row(row)

def coal_type_refine(s):
    global coal_type_string
    if type(s) == type(float()) or NOT_SPACE.search(str(s)) == None:
        s = coal_type_string.strip()
    if s.startswith("А"):
        s = "антрацит"
    elif s.startswith("Г"):
        s = "газове"
    elif s.startswith("П"):
        s = "пісне"
    return s
    
def is_stopped(row):
    return sum([str(r.value).lower().strip().startswith("зупинена") for r in row]) > 0



def parse_row(row):
    global region, enterprise, is_enterprise, f, anomalia, date, coal_type_string
    global df_stations
    global reserve_id, spending_id
    if reserve_id:
        if not is_blank(row[0]):
            if row[0].value != row[0].value.upper():
                station = row[0].value.replace("Ө", "").strip()
                if "вугілля" in station.lower():
                    coal_type_string = station
                reserve = row[reserve_id].value
                if (not reserve) and reserve != 0:
                    reserve = None
                else:
                    reserve = round(float(row[reserve_id].value))
                max_coal = row[reserve_id-8].value
                min_coal = row[reserve_id-7].value
                if row[reserve_id-1].value != "":
                    plan_coal = round(float(row[reserve_id-1].value))
                else:
                    plan_coal = None
                spending = row[spending_id+1].value
                if spending != "" and spending != 0:
                    days = round(reserve / spending)
                elif spending == 0:
                    days = 0
                else:
                    days = None
                """if (spending == 0 and reserve == 0) or is_stopped(row):
                    stopped = True
                else:
                    stopped = False"""
                stopped = is_stopped(row)
                delivery = row[spending_id - 1].value
                coal_type = coal_type_refine(row[1].value)
                try:
                    plan_percent = round(float(reserve) / float(plan_coal) * 100)
                    completance = round(float(reserve) - float(plan_coal))
                except:
                    plan_percent = None
                    completance = None
                df_temp = pd.DataFrame([[station, date, reserve,  min_coal, max_coal, plan_coal, plan_percent, completance, coal_type, delivery, spending, days, stopped, f]], columns = CSV_HEADERS)
                df_stations = df_stations.append(df_temp, ignore_index = True)   
    else:
        if not is_blank(row[0]):
            for i in range(len(row)):
                if str(row[i].value).strip().startswith("Запас"):
                    reserve_id = i
                if str(row[i].value).strip().startswith("Витрата"):
                    spending_id = i
     
def is_completely_stopped(group):
    not_working = sum(group['stopped'])
    coal_types_all = len(group['stopped'])
    return  not_working == coal_types_all


warnings.filterwarnings('ignore')

df_ids = pd.read_csv(STATIONS_IDS_FILE)
stations_dict = {}
for i in range(df_ids.shape[0]):
    stations_dict[df_ids.loc[i, 'station_title_original']] = {'id': df_ids.loc[i,'station_id'], "short":df_ids.loc[i,'station_title_short']} 

walk = os.walk(INPUT_DATA_FOLDER)
full_filenames = []
for dirpath, subdirs, files in walk:
    for f in files:
        full_filenames.append(os.path.join(dirpath, f))
form = [f for f in full_filenames if file_re.search(f)]
rows = []
anomalia = []
df_stations = pd.DataFrame(columns = CSV_HEADERS)


for f in form:
    region = None
    enterprise = None
    is_enterprise = None
    sheet = None
    wb = None
    year_month = None
    mine = None
    get_workbook(f)

with open(STATIONS_INFO_FILE, "r") as sif:
    stations_info_dict = json.load(sif)



static_info_keys = [h.replace("_jan2016", "") for h in OUTPUT_STATIC_HEADERS[1:]]
static_info_rows = []
for k in stations_info_dict:
    row = [stations_info_dict[k][r] for r in static_info_keys]
    row = [k] + row
    static_info_rows.append(row)
df_stations_static = pd.DataFrame(static_info_rows, columns = OUTPUT_STATIC_HEADERS)

df_stations['station'] = df_stations['station'].str.replace("\n", "")
df_stations = df_stations.loc[df_stations['station'].isin(df_ids['station_title_original']), :]
df_stations['date'] = pd.to_datetime(df_stations['date'], format = '%Y-%m-%d')
df_stations['id'] = df_stations['station'].map(lambda s: stations_dict[s]['id'])
df_stations['station'] = df_stations['station'].map(lambda s: stations_dict[s]['short'])
df_stations = df_stations.loc[:, ["id", "station", "date", "reserve", "min", "max", "planned", "plan_percent", 'completance', "coal_type", "delivery", "spending", "days"]]
df_stations_output = df_stations.merge(df_stations_static, left_on = "id", right_on = "station_id")
df_stations_output = df_stations_output.loc[:, df_stations_output.columns != "id"]
df_stations_output.to_csv(OUTPUT_FILE, index = False)

for i in df_stations.index:
    if df_stations.loc[i, 'coal_type'] == "газове":
        df_stations['id'][i] += "_g"
    elif df_stations.loc[i, 'coal_type'] == "антрацит":
        df_stations['id'][i] += "_a"
    elif df_stations.loc[i, 'coal_type'] == "пісне":
        df_stations['id'][i] += "_p"



stations_json = []
stations = df_stations['id'].drop_duplicates()
for st in stations:
    station = df_stations.loc[df_stations['id'] == st, :]
    station['date'] = station['date'].map(lambda x: str(x.date()))
    station_dict = {}
    station_dict['station'] = station['station'].values[0]
    station_dict['id'] = station['id'].values[0]
    station_dict.update(stations_info_dict[station_dict['id'].split("_")[0]])
    station_dict['mentions'] = []
    for i in station.index:
        station_dict['mentions'].append(station.loc[i, ["date", "reserve", "min", "max", "planned", "plan_percent", "coal_type", "days"]].to_dict())
    if len(station_dict['mentions']) > 1:
        stations_json.append(station_dict)

df_stations30 = df_stations.loc[df_stations['date'] >= max(df_stations['date'] + datetime.timedelta(-30)), :]     
df_stations = df_stations.loc[df_stations['date'] == max(df_stations['date']), :]

for i in df_stations.index:
    if df_stations.loc[i, 'id'].endswith("_g"):
        id_a = df_stations.loc[i, 'id'].replace('_g', "_a")
        if df_stations[df_stations['id'] == id_a].shape[0] > 0:
            df_stations['plan_percent'][i] =  df_stations['plan_percent'][df_stations['id'] == id_a].values[0]
        else:
            id_p = df_stations.loc[i, 'id'].replace('_g', "_p")
            if df_stations[df_stations['id'] == id_p].shape[0] > 0:
                df_stations['plan_percent'][i] =  df_stations['plan_percent'][df_stations['id'] == id_p].values[0]

df_stations['date'] = df_stations['date'].map(create_date_string)

stations_json_final = []
for st_d in stations_json:
    station = df_stations.loc[df_stations['id'] == st_d['id'], :]
    station30 = df_stations30.loc[df_stations30['id'] == st_d['id'], :]
    postfix = "_" + st_d['id'].split('_')[1]
    st_d['min' + postfix] = blank_string_to_null(station['min'].values[0])
    st_d['max' + postfix] = blank_string_to_null(station['max'].values[0])
    st_d['last30days_delivery' + postfix] = round(station30['delivery'].sum(), 2)
    st_d['last30days_spending' + postfix] = round(station30['spending'].sum(), 2)
    st_d['reserve' + postfix] = station['reserve'].values[0]
    for p in STATIONS_POSTFIX:
        if p != postfix:
            st_other_type = df_stations.loc[df_stations['id'] == st_d['id'].split("_")[0] + p, :]
            st_other_type30 = df_stations30.loc[df_stations30['id'] == st_d['id'].split("_")[0] + p, :]
            if st_other_type.shape[0] > 0:
                st_d['min' + p] = blank_string_to_null(st_other_type['min'].values[0])
                st_d['max' + p] = blank_string_to_null(st_other_type['max'].values[0])
                st_d['reserve' + p] = blank_string_to_null(st_other_type['reserve'].values[0])
                st_d['last30days_delivery' + p] = round(st_other_type30['delivery'].sum(), 2)
                st_d['last30days_spending' + p] = round(st_other_type30['spending'].sum(), 2)
            else:
                st_d['min' + p] = -1
                st_d['max' + p] = -1
                st_d['reserve' + p] = -1
                st_d['last30days_delivery' + p] = -1
                st_d['last30days_spending' + p] = -1
    st_d['completance'] = station['completance'].values[0]
    st_d['plan_percent'] = station['plan_percent'].values[0]
    st_d['date'] = station['date'].values[0]

if not os.path.exists(OUTPUT_DATA_FOLDER):
    os.makedirs(OUTPUT_DATA_FOLDER)

with open(os.path.join(OUTPUT_DATA_FOLDER, JSON_DATA_FILE), "w") as jf:
    json.dump(stations_json, jf, ensure_ascii = False, indent = 4)


