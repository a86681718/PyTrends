import xlrd
from pytrends.pyGTrends import pyGTrends
import time
import csv
from random import randint
__author__ = 'fang'

# Get ticker list
path = "index_const.xlsx"
book = xlrd.open_workbook(path)
sp1500 = book.sheet_by_index(1)
tickers = sp1500.col_values(12)

# GoogleTrend api settings
google_username = "b86681718@gmail.com"
google_password = "Fang8426ainos"
path = ""
connector = pyGTrends(google_username, google_password)

# Datetime range
date = []
for i in range(2011,2016):
    d = 1
    for j in range(1,5):
        if d<10:
            date.append("0%d/%d 3m" % (d, i))
        else:
            date.append("%d/%d 3m" % (d, i))
        d += 3


def save_csv(trend_name, result_data):
    filename = trend_name + ".csv"
    with open(filename, mode='wb') as f:
        f.write(result_data.encode('utf8'))

count = 0
# Get tickers search result
for tic in tickers:
    result = ''
    count += 1
    if count <= 71:
        continue
    for d in date:
        # Make request
        connector.request_report(tic, 'en-US', None, None, d, None)
        # Wait a random amount of time between requests to avoid bot detection
        #          time.sleep(randint(1, 3))
        # Extract result string
        part_result = connector.decode_data
        start = part_result.find('Day,')+5+len(tic)
        end = part_result.find('Top regions')-2
        result += part_result[start:end]
    save_csv(tic, result)



