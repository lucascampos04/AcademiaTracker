from openpyxl.styles import PatternFill, Font
import openpyxl
from tkinter import Tk, simpledialog
from datetime import datetime
import os

file = "./Academia.xlsx"

if os.path.exists(file):
    wb = openpyxl.load_workbook(file)
    sheet = wb.active
else:
    wb = openpyxl.Workbook()
    sheet = wb.active

    sheet["A1"] = 'Exercicio'
    sheet["B1"] = 'Carga'
    sheet["C1"] = 'Repetições'
    sheet["D1"] = 'Data'

    header_style = PatternFill(start_color="B0C4DE", end_color="B0C4DE", fill_type="solid") 
    font_style = Font(bold=True)

    for col_num in range(1, 5):
        sheet.cell(row=1, column=col_num).fill = header_style
        sheet.cell(row=1, column=col_num).font = font_style

def inserirDados():
    exercicio = simpledialog.askstring("Exercício", "Digite o exercício")
    carga = simpledialog.askinteger("Carga", "Digite a carga")
    repeticoes = simpledialog.askinteger("Repetições", "Digite a quantidade de repetições")
    data = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    nova_linha = (exercicio, carga, repeticoes, data)
    sheet.append(nova_linha)

    print(f'Exercício: {exercicio}, Carga: {carga}, Repetições: {repeticoes}, Data: {data}')

inserirDados()

wb.save(file)
