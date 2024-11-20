import json
ev_dict = {
    'Йо.':'iv',
    'Лк.':'lk',
    'Мт.':'mt',
    'Мр.':'mr'
    }

def split_words(text):
    words = []
    current_word = ""
    for char in text:
        if char.isalnum():
            current_word += char
        else:
            if current_word:
                words.append(current_word)
                current_word = ""
    if current_word:
        words.append(current_word)
    return words

def convert_apostol_item(item, reading_list):
    item_words = split_words(item[6:])
    for r in reading_list:
        reading_words = split_words(r['reading'])
        if reading_words == item_words:
            zach_filename = f"ap{r['zach_num']:0>3}{r['zach_abc']}"
            a_string=f"<a href='#' data-remote='../zachala/{zach_filename}.html' data-read='../benedict/{zach_filename}.html' data-toggle='modal' data-target='#exampleModal' >{item}</a>"
            return a_string
    return item

def convert_evangelie_item(item, reading_list, ev_dict):
    item_words = split_words(item[6:])
    print(item_words)
    book_abbrev = item_words[0]+'.'
    if book_abbrev in ev_dict:
        book_prefix = ev_dict[book_abbrev]
        for r in reading_list:
            if split_words(r['reading']) == item_words:
                zach_filename = f"{book_prefix}{r['zach_num']:0>3}{r['zach_abc']}"
                a_string=f"<a href='#' data-remote='../zachala/{zach_filename}.html' data-read='../benedict/{zach_filename}.html' data-toggle='modal' data-target='#exampleModal' >{item}</a>"
                return a_string
    return item

input_string='<i>Великі (Царські) часи.</i><br>На 1-му часі<br>Ап. – Ді. 33 зач.; 13, 25-33.<br>Єв. – Мт. 5 зач.; 3, 1-11.<br>На 3-му часі<br>Ап. – Ді. 42 зач.; 19, 1-8.<br>Єв. – Мр. 1 зач.; 1, 1-8.<br>На 6-му часі<br>Ап. – Рим. 91 зач.; 6, 3-11.<br>Єв. – Мр. 2 зач.; 1, 9-15.<br>На 9-му часі<br>Ап. – Тит. 302 зач.; 2, 11-14; 3, 4-7.<br>Єв. – Мт. 6 зач.; 3, 13-17.<br><i>Літургія св. Василія Великого з вечірнею.</i><br>Ап. – 1 Кор. 143 зач.; 9, 19-27.<br>Єв. – Лк. 9 зач.; 3, 1-18.<br><i>На осв. води:</i><br>Ап. – 1 Кор. 143 зач. (від половини); 10, 1-4.<br>Єв. – Мр. 2 зач.; 1, 9-11.'

with open('apostol.json', 'r', encoding='utf-8') as j:
    apostol_reading_list = json.loads(j.read())

for item_start in input_string.split("Ап. – "):
    if item_start:
        item = "Ап. – " + item_start.split("<br>")[0]
        converted_item = convert_apostol_item(item, apostol_reading_list)
        input_string = input_string.replace(item, converted_item)

with open('evanhelie.json', 'r', encoding='utf-8') as j:
    evangelie_reading_list = json.loads(j.read())

for item_start in input_string.split("Єв. – "):
    if item_start:
        item = "Єв. – " + item_start.split("<br>")[0]
        converted_item = convert_evangelie_item(item, evangelie_reading_list, ev_dict)
        input_string = input_string.replace(item, converted_item)

print(input_string)
