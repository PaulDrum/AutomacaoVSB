from datetime import date, timedelta

yesterday = date.today() - timedelta(days=1)
d1 = yesterday.strftime("%d.%m.%Y")
