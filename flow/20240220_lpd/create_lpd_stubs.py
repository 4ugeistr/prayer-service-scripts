import re, easygui, calendar, glob,shutil,os
from datetime import datetime
import paschalia

month_no = datetime.now().month+1 if datetime.now().month!=12 else 1
year_no = datetime.now().year if datetime.now().month!=12 else datetime.now().year+1

mode = easygui.choicebox('u - Юліанський, g - Григоріанський', 'Вибір календаря', ['u','g'])

day_short_dic={"ПН":1,
        "ВТ":2,
        "СР":3,
        "ЧТ":4,
        "ПТ":5,
        "СБ":6,
        "НД":7}
day_short_dic_reversed = {v:k for k,v in day_short_dic.items()}


lpd_templates = glob.glob('ЛПД/*/*.docx')
#lpd_templates_filenames = [x.split('\\')[-1] for x in lpd_templates]

if __name__ == "__main__":
    paschalia_dates = paschalia.get_prev_next_pascha(datetime(year_no, month_no,1), mode)
    stub_dic={}
    folder= f'drafts\\{year_no}-{month_no:02}-{mode}'
    #os.makedirs('drafts', exist_ok=True)
    os.makedirs(folder, exist_ok=True)

    for d in range(1,calendar.monthrange(year_no, month_no)[1]+1):
        #stub_dic[d]=None
        day_details = paschalia.get_day_details(datetime(year_no,month_no,d),paschalia_dates)
        expected_template_path = f"ЛПД\\тиждень-{day_details[1]}\\{day_details[1]}-{day_details[3]}-ЛПД.docx"
        
        if expected_template_path in lpd_templates:
            filename=folder+f'\\{d:02}-{day_details[1]}-{day_short_dic_reversed[day_details[3]].lower()}-ЛПД.docx'
            shutil.copy2(expected_template_path, filename)
            stub_dic[d]=filename
        #re_result = re.search("(\d)-(\d)-ЛПД.docx")
    print(f"Created {len(stub_dic)} stubs for month {month_no} mode {mode}!")