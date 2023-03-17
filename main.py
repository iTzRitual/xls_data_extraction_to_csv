import os
import tkinter as tk
import xlrd
from datetime import datetime
from tkinter import messagebox

TO_IMPORT = 'import_xls'
TO_EXPORT = 'exported_csv'


def main():
    init_tkinter()

    if not os.path.exists(TO_IMPORT):
        os.makedirs(TO_IMPORT)

    if not os.path.exists(TO_EXPORT):
        os.makedirs(TO_EXPORT)

    isDoneAnything = False

    for file in os.listdir(TO_IMPORT):
        if file.endswith('.xls'):
            workbook = xlrd.open_workbook(os.path.join(TO_IMPORT, file))
            print("The number of worksheets is {0}".format(workbook.nsheets))
            for sheet in workbook.sheets():
                cell_positions = []
                array_for_rows = []
                for row_index, row in enumerate(range(sheet.nrows)):
                    if row_index == 0:
                        for cell_index, cell in enumerate(sheet.row(row_index)):
                            if sheet.cell(row_index, cell_index).value == "Nazwa" or sheet.cell(row_index,
                                                                                                  cell_index).value == "Kwota" or sheet.cell(
                                    row_index, cell_index).value == "Data płatności" or sheet.cell(row_index,
                                                                                                  cell_index).value == "Faktura" or sheet.cell(
                                    row_index, cell_index).value == "Kupujący":
                                cell_positions.append([cell_index, sheet.cell(row_index, cell_index).value])
                    else:
                        row_value = ""
                        for cell_index, cell in enumerate(sheet.row(row_index)):
                            for i in cell_positions:
                                if cell_index == i[0]:
                                    if i[1] == "Nazwa" or i[1] == "Kupujący":
                                        row_value += "\"" + str(sheet.cell(row_index, cell_index).value) + "\"" + ";"
                                    elif i[1] == "Data płatności":
                                        date = datetime(*xlrd.xldate_as_tuple(sheet.cell(row_index, cell_index).value, workbook.datemode))
                                        row_value += date.strftime("%d/%m/%Y") + ";"
                                    else:
                                        row_value += str(sheet.cell(row_index, cell_index).value) + ";"
                        row_value = row_value[:-1]
                        array_for_rows.append(row_value)
            file = file[:-4] + ".csv"
            if os.path.isfile(os.path.join(TO_EXPORT, file)):
                i = 1
                while os.path.isfile(os.path.join(TO_EXPORT, f"{i}_{file}")):
                    i += 1
                file = f"{i}_{file}"

            with open(os.path.join(TO_EXPORT, file), 'w') as f:
                f.write("Nazwa;Kwota;Data płatności;Faktura;Kupujący\n")
                for row in array_for_rows:
                    f.write(row + "\n")

            isDoneAnything = True

    if isDoneAnything:
        messagebox.showinfo("Success", "Done")
    else:
        messagebox.showinfo("Error", "Put .xls files in import_xls folder")


def init_tkinter():
    root = tk.Tk()
    root.withdraw()
    root.silence_deprecation = True
    return root


if __name__ == '__main__':
    main()
