from pytrends.pyGTrends import pyGTrends
import time
from random import randint

google_username = "a86681718@gmail.com"
google_password = "amy8426ainos"
path = ""

# connect to Google
connector = pyGTrends(google_username, google_password)

# make request
connector.request_report("Pizza")

# wait a random amount of time between requests to avoid bot detection
time.sleep(randint(5, 10))

# download file
connector.save_csv(path, "pizza")

# get suggestions for keywords
keyword = "milk substitute"
data = connector.get_suggestions(keyword)
print(data)

