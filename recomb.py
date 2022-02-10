from numpy import mat
import pandas as pd
import openpyxl 
import xlsxwriter

def OpenPlan(filename_plan):
    wb = openpyxl.load_workbook(filename_plan)
    ws = wb.active
    return  ws

def reStructure(filename_plan):
    OpenPlan(filename_plan)
    workbook = openpyxl.load_workbook(filename_plan)
    worksheet = workbook.active
    for i in range(1, 400):
        match worksheet["E" + str(i)].value:
            case "Первый семестр":
                worksheet["E" + str(i)].value = "1 семестр"
                int(i)
            case "Второй семестр":
                worksheet["E" + str(i)].value = "2 семестр"
                int(i)
            case "Третий семестр":
                worksheet["E" + str(i)].value = "3 семестр"
                int(i)
            case "Четвертый семестр":
                worksheet["E" + str(i)].value = "4 семестр"
                int(i)
            case "Пятый семестр":
                worksheet["E" + str(i)].value = "5 семестр"
                int(i)
            case "Шестой семестр":
                worksheet["E" + str(i)].value = "6 семестр"
                int(i)
            case "Седьмой семестр":
                worksheet["E" + str(i)].value = "7 семестр"
                int(i)
            case "Восьмой семестр":
                worksheet["E" + str(i)].value = "8 семестр"
                int(i)
            case "Девятый семестр":
                worksheet["E" + str(i)].value = "9 семестр"
                int(i)
            case "Десятый семестр":
                worksheet["E" + str(i)].value = "10 семестр"
                int(i)
            case "Одинадцатый семестр":
                worksheet["E" + str(i)].value = "11 семестр"
                int(i)

    workbook.save(filename=filename_plan)

        

filename_plan = '1c_Data.xlsx'
reStructure(filename_plan)


