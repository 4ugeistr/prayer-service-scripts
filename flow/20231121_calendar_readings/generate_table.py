import calendar,csv
import paschalia
from datetime import datetime

year_from = 2021
year_to = 2024

if __name__ == "__main__":
    rows=[]
    for year in range(year_to, year_from, -1):
        print("Calculating pascha dates for", datetime(year,1,1))
        paschalia_dates = paschalia.get_prev_next_pascha(datetime(year,1,1))
        for month in range(1,7):
            for day in range(1,calendar.monthrange(year, month)[1]+1):
                rows.append([f"{year}.{month:02}.{day:02}"]+list(paschalia.get_day_details(datetime(year,month,day),paschalia_dates)))


    with open(f'paschalia_{year_from}-{year_to}.csv','w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter=',', quotechar='"')
        csvwriter.writerows(rows)
    print("DONE!")