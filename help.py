from datetime import datetime
from datetime import date
import calendar

currentMonth = datetime.now().month
currentYear = datetime.now().year
_,num_days = calendar.monthrange(2016, 3)
first_day = date(currentYear, currentMonth, 1)
last_day = date(currentYear, currentMonth, num_days)
print(first_day.strftime('%Y-%m-%d'))
print(last_day.strftime('%Y-%m-%d'))