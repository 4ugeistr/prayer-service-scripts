import docx2txt
import re
import json
import docx

# Read the contents of the docx file
text = docx2txt.process("09-ТРАВЕНЬ.docx")
doc = docx.Document('09-ТРАВЕНЬ.docx')


troparia ={'voskresni':{},
 'menaion':{5:{}},
 'triodion':{}}

for p in doc.paragraphs:
    re_result=re.search('(\d{1,2}) \((\d{1,2})\)',p.text)
    if re_result:
        print(re_result)
        print(re_result.groups())
        troparia['menaion'][5][re_result.group(2)]={}
        troparia['menaion'][5][re_result.group(2)]['ju']=re_result.group(1)
        troparia['menaion'][5][re_result.group(2)]['gr']=re_result.group(2)


# Split the text into sections based on the section headers
sections = re.split(r"\n\d{1,2} \((\d{1,2})\) ", text)[1:]

# Parse the sections into a dictionary format
parsed_sections = []
for section in sections:
    lines = section.strip().split("\n")
    section_data = {}
    section_data["title"] = lines[0]
    section_data["content"] = "\n".join(lines[1:])
    parsed_sections.append(section_data)

# Write the parsed sections to a JSON file
with open("test.json", "w",encoding='utf8') as f:
    json.dump(parsed_sections, f,ensure_ascii=False)
