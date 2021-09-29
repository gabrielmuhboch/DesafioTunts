#pip3 install openpyxl
from openpyxl import Workbook, load_workbook;

sheet = load_workbook(r'C:\Users\Microsoft\Desktop\Desafio.xlsx');
active = sheet.active;
totalClasses = 60;

##setting a lopp so i can work with each student
for cell in active["A"]:

    #getting the active row
    row = cell.row;

    #setting the final's test grade to be 0 by default so we change it only if the student nedds to
    active[f"H{row}"] = "0"

    #checking if the active row is the head of the spreadsheet
    if cell.value == "Matricula":
        continue
    #Firstly, we need to verify if the student has enought classe's attendence not to fail
    elif active[f"C{row}"].value/totalClasses > 0.25:
        #if the student failed, need to writ it on the new spreadsheet
        active[f"G{row}"] = "Reprovado por Falta"

    #
    else:
        #getting the student's tests average
        sum = active[f"D{row}"].value + active[f"E{row}"].value + active[f"F{row}"].value;
        avarege = sum / 30;

        #verify if the student avarege isn't enought for the final exam
        if avarege < 5:
            active[f"G{row}"] = "Reprovado por Nota"

        #verify if the student will need the final exam
        elif 5<= avarege < 7:
            active[f"G{row}"] = "Exame Final"

            #getting the student's minimum score he needs on the final test
            naf = 10 - avarege
            naf = format(naf, ".1f")
            active[f"H{row}"] = naf;

        #verify if the student passed
        elif avarege >= 7:
            active[f"G{row}"] = "Aprovado"

#saving the new spreedsheet
sheet.save(r'C:\Users\Microsoft\Desktop\Engenharia de Software – Desafio [Gabriel Muhlstedt Bochnia].xlsx')
