import openpyxl
from docxtpl import DocxTemplate
import datetime

# Load data from excel file
path = "data.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.active
list_values = list(sheet.values)
# print(list_values)
transcript = DocxTemplate("template.docx")

now = datetime.datetime.now()
currentDate = now.strftime("%d %B %Y")

for value in list_values[1:]:
   print(value)
   transcript.render({
    "student_id": value[0],
    "first_name": value[1],
    "last_name": value[2],
    "logic": value[3],
    "l_g": value[4],
    "bcum": value[5],
    "bc_g": value[6],
    "design": value[7],
    "d_g": value[8],
    "p1": value[9],
    "p1_g": value[10],
    "e1": value[11],
    "e1_g": value[12],
    "wd": value[13],
    "wd_g": value[14],
    "algo": value[15],
    "al_g": value[16],
    "p2": value[17],
    "p2_g": value[18],
    "e2": value[19],
    "e2_g": value[20],
    "sd": value[21],
    "sd_g": value[22],
    "js": value[23],
    "js_g": value[24],
    "php": value[25],
    "ph_g": value[26],
    "db": value[27],
    "db_g": value[28],
    "vc1": value[29],
    "v1_g": value[30],
    "node": value[31],
    "no_g": value[32],
    "e3": value[33],
    "e3_g": value[34],
    "p3": value[35],
    "p3_g": value[36],
    "oop": value[37],
    "op_g": value[38],
    "lar": value[39],
    "lar_g": value[40],
    "vue": value[41],
    "vu_g": value[42],
    "vc2": value[43],
    "v2_g": value[44],
    "e4": value[45],
    "e4_g": value[46],
    "p4": value[47],
    "p4_g": value[48],
    "int": value[49],
    "in_g": value[50],
    "cur_date": currentDate

   })
   doc_name =  str(value[1]) + "-" + str(value[2]) + ".docx"
   transcript.save(doc_name) 