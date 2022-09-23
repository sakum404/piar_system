from datetime import datetime


# date = datetime(datetime.today().year, datetime.today().month, datetime.today().day)
date = datetime.strftime(datetime.today(), "%b %d %Y")

print(date)