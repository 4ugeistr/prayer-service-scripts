from datetime import datetime, timedelta

def get_echos(current_date):
    return (int(datetime(2023, 9, 5).strftime("%U"))-25) % 8 + 1

print(get_echos(datetime(2023, 9, 5)))





