from pyquery import PyQuery as pq
import re 
import pandas as pd
from time import sleep
import os
import subprocess
from urllib.request import urlretrieve
from datetime import date, datetime

INPUT_DATA_FOLDER = "../coal_input"
COAL_URL = "http://mpe.kmu.gov.ua/minugol/control/uk/publish/officialcategory?cat_id=245183254"
COAL_SELECTOR = '.text_news a'
DOWNLOAD_SELECTOR = '.MsoNormal a'
DATE_RE = re.compile("\d{2}\.\d{2}\.\d{4}")
DF_FILE = os.path.join(INPUT_DATA_FOLDER, 'coal_reserves_stations.csv')
BASIC_URL = 'http://mpe.kmu.gov.ua/minugol/'
SLEEP_TIME = 1


REFRESHMENT_PERIOD = 60 * 60


def change_date_format(s):
    splt = s.split('.')
    return '-'.join([splt[2], splt[1], splt[0]])

def check_latest():
    page = pq(COAL_URL)
    days_links = page(COAL_SELECTOR)
    days_links = [d for d in days_links if DATE_RE.search(d.text)]
    dates = [change_date_format(DATE_RE.search(d.text).group(0)) for d in days_links]
    filenames = [d.text for d in days_links]
    hrefs = [pq(d).attr('href') for d in days_links]
    linksWithDates = zip(dates, filenames, hrefs)
    thereWasNewDates = False
    for i in linksWithDates:
        if pd.to_datetime(i[0]) > latest:
            try:
                download_file(i)
            except:
                print("Some problems with dowload, will try again in {s:d} seconds".format(s = REFRESHMENT_PERIOD))
                thereWasNewDates = False
                break
            thereWasNewDates = True
    if thereWasNewDates:
        import coal_reserves_daily
    
def download_file(f):
    sleep(SLEEP_TIME)
    link = BASIC_URL + f[2]
    filename = os.path.join(INPUT_DATA_FOLDER, f[1] + '.xlsx')
    page = pq(link)
    download_link = page(DOWNLOAD_SELECTOR).attr('href')
    #print(len(download_link))
    urlretrieve(download_link, filename)
    

df = pd.read_csv(DF_FILE)
df['date'] = pd.to_datetime(df['date'])
latest = max(df['date'])
today = datetime.now()
today = today.replace(hour=0, minute=0, second=0, microsecond=0)

while True:
    print(latest)
    if latest <  today:
        check_latest()
        deploy_results()
    sleep(REFRESHMENT_PERIOD)