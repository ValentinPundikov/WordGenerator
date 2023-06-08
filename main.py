from docxtpl import DocxTemplate
from num2words import num2words
import csv
import random
from docx2pdf import convert


file_name = "C:/Users/Админ/PycharmProjects/WordGenerator/data1.txt"
doc = DocxTemplate("C:/Users/Админ/PycharmProjects/WordGenerator/template.docx")

with open(file_name, "r", encoding="utf-8") as file:
    reader = csv.reader(file)
    for row in reader:
        rub = str(num2words(row[5], lang='ru'))
        rub_str = rub + " руб."
        kop = str(num2words(row[6], lang='ru'))
        kop_str = kop + " копеек"
        res = str(rub + " руб. " + kop + " копеек ")
        context = {"var_name1": row[0], "var_name2": row[1], "var_name3": row[2], "var_name4": row[3]
            ,"var_name5": row[4], "var_name6": rub_str, "var_name7": kop_str, "var_name8": row[5], "var_name9": row[6]}
        doc.render(context)
        rand = str(random.randint(1,27))
        filename = str(row[0] + ".docx")
        doc.render(context)
        doc.save(filename)
        pdf = str(row[0] + ".pdf")
        convert(filename)
