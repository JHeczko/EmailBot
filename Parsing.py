from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font, Alignment, PatternFill


def edit_excel(workbook):
    newWorkBook = Workbook()
    newSheet = newWorkBook.active
    sheet = workbook.active  # Pobieramy aktywny arkusz

    # tablice oparte o Imie oraz nazwiska mamy
    hashImieR1 = {}
    hashImieR2 = {}
    hashImieR3 = {}
    hashImieALL = {}
    hashImieMail = {}
    hashImieIlosc = {}

    try:
        # =-=-=-=-=-=-=-=-=SORTING DATA FROM OLD WORKBOOK=-=-=-=-=-=-=-=-=
        for row in sheet.iter_rows(min_row=2, min_col=3, max_col=11, values_only=True):
            #dzieckoNazwisko = row[0].split()[1] if row[0] is not None else 'nieznane dziecko'
            mama = row[2] if row[2] is not None else 'nieznany_rodzic'
            #tataNazwisko = row[3] if row[3] is not None else 'nieznany rodzic'
            mail = row[4] if row[4] is not None else ''
            moneyR1 = row[6] if row[6] is not None else 0
            moneyR2 = row[7] if row[7] is not None else 0
            moneyR3 = row[8] if row[8] is not None else 0
            moneyAll = moneyR1 + moneyR2 + moneyR3

            if mama in hashImieR1:
                hashImieR1[mama] += moneyR1
            else:
                hashImieR1[mama] = moneyR1

            if mama in hashImieR2:
                hashImieR2[mama] += moneyR2
            else:
                hashImieR2[mama] = moneyR2

            if mama in hashImieR3:
                hashImieR3[mama] += moneyR3
            else:
                hashImieR3[mama] = moneyR3

            if mama in hashImieALL:
                hashImieALL[mama] += moneyAll
            else:
                hashImieALL[mama] = moneyAll

            if mama in hashImieIlosc:
                hashImieIlosc[mama] += 1
            else:
                hashImieIlosc[mama] = 1

            hashImieMail[mama] = mail
    except Exception as e:
        print(f"Data sep loop error: {e}")

    try:
        # =-=-=-=-=-=-=-=-=STYLING AND FILLING NEW WORKBOOK=-=-=-=-=-=-=-=-=

        # ============Styling============
        header_font = Font(bold=True)
        header_alignment = Alignment(horizontal='center', vertical='center')
        cell_alignment = Alignment(horizontal='left', vertical='center')
        light_gray_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')


        # ============Header creation============
        headers = ['Zwrot','Nazwisko', 'Mail', 'Ilość Dzieci', 'R1', 'R2', 'R3', 'All']
        for col_num, header in enumerate(headers, start=1):
            cell = newSheet.cell(row=1, column=col_num, value=header)
            cell.font = header_font
            cell.alignment = header_alignment
            cell.fill = PatternFill(start_color='00969696', end_color='00969696', fill_type='solid')

        # ============Filling data============
        row = 2
        for mama in hashImieMail.keys():
            nazwiskoMama = mama.split()[1:]
            nazwiskoMama = ''.join(nazwiskoMama)
            newSheet.cell(row=row, column=1, value="Szanowna Pani")
            newSheet.cell(row=row, column=2, value=nazwiskoMama if mama!="nieznany_rodzic" else "Nie znaleziono mamy!")
            newSheet.cell(row=row, column=3, value=hashImieMail[mama]if mama!="nieznany_rodzic" else "Uzupelnij dane o imie oraz nazwisko mamy")
            newSheet.cell(row=row, column=4, value=hashImieIlosc[mama])
            newSheet.cell(row=row, column=5, value=hashImieR1[mama]).number_format = '#,##0.00 "zł"'
            newSheet.cell(row=row, column=6, value=hashImieR2[mama]).number_format = '#,##0.00 "zł"'
            newSheet.cell(row=row, column=7, value=hashImieR3[mama]).number_format = '#,##0.00 "zł"'
            newSheet.cell(row=row, column=8, value=hashImieALL[mama]).number_format = '#,##0.00 "zł"'
            row += 1

        # ============Styling============
        row = 2
        for mama in hashImieMail.keys():
            for i in range(1,9):
                if row % 2 != 0:
                    newSheet.cell(row=row, column=i).fill = light_gray_fill
            row += 1

        row = 2
        for mama in hashImieMail.keys():
            for i in range(1,9):
                newSheet.cell(row=row, column=i).alignment = cell_alignment
            row += 1

        # Adjust column widths
        for col in newSheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except Exception:
                    pass
            adjusted_width = (max_length + 10)  # Add some extra space
            newSheet.column_dimensions[column].width = adjusted_width
    except Exception as e:
        print(f"Styling loop error: {e}")
    return newWorkBook

if __name__ == "__main__":
    workbook = load_workbook("./tests/test.xlsx")
    edited_workbook = edit_excel(workbook)
    edited_workbook.save("./testy_parsing.xlsx")