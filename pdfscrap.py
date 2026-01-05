import pdfplumber
import pandas as pd
# openpyxl
import os
from contextlib import redirect_stderr
import easyocr
from PIL import Image
import numpy as np

def createErrorMsg():
    msg = f"FileName: {fileName} has failed."

    if not foundDate:
        msg += f"\nInvoice Data: {invoiceDate}."

    if not foundInvoiceNum:
        msg += f"\nInvoice Num: {invoiceNum}."

    if not foundDate or not foundInvoiceNum:
        msg += f"\nFirstPageText\n{firstPageText}"

    return msg

def getTables(firstPage, lastPage):
    for pageNum in range(firstPage - 1, lastPage):  # checks the selected pages
        table = pdf.pages[pageNum].extract_table()

        if table is not None:
            df = pd.DataFrame(table[1:], columns=table[0]).replace(["None", ""], pd.NA)

            if not df.isna().all().all():
                tables.append(pd.DataFrame(table[1:], columns=table[0]))


# CHANGE
blankLineCount = 0  # amount of blank lines between each pdf
wordAfterVAT = "Total"
columnWithVAT = 9

# DON'T CHANGE
invoiceData = []; errorMsgs = []
fileCount = 1; vatRowCount = 0; vatColCount = 0; errorsRowCount = 0
currentDir = os.path.dirname(os.path.abspath(__file__))

if not os.path.exists(currentDir + "\\pdfs"):  # checks the folder pdfs exists
    os.mkdir("pdfs")

files = os.listdir(currentDir + "\\pdfs")

if len(files) == 0:  # handle no pdfs
    print("No pdfs found")
    quit()

with open(os.devnull, "w") as fNull:  # remove error msgs
    with redirect_stderr(fNull):  # remove pdfplumber error msgs
        for fileName in files:
            print(f"File {fileCount}/{len(files)}")  # progress update
            fileCount+=1

            invoiceDate = f"Not Found"; invoiceNum = f"Not Found"
            tables = []
            foundDate = False; foundInvoiceNum = False

            if ".pdf" in fileName:
                    with pdfplumber.open("pdfs\\" + fileName) as pdf:
                        firstPageText = pdf.pages[0].extract_text().split("\n")

                        try:
                            for textRow in firstPageText:  # loops through each row of text in the first page
                                textRowSpace = textRow.split(" ")  # split into spaces to try and just the date / num

                                if "DATE" in textRow.upper() or "TAX POINT" in textRow.upper():  # get date row text
                                    if len(textRowSpace) > 3:
                                        for textSpace in textRowSpace:
                                            if textSpace[0].isdigit():
                                                invoiceDate = f"{textSpace} {textRowSpace[textRowSpace.index(textSpace)+1]} {textRowSpace[textRowSpace.index(textSpace)+2]}"
                                                break
                                    else:
                                        invoiceDate = textRow
                                    foundDate = True

                                if "INVOICE NO" in textRow.upper():  # get invoice num row text
                                    if len(textRowSpace) > 2:
                                        for textSpace in textRowSpace:
                                            if textSpace[0].isdigit():
                                                invoiceNum = textSpace

                                    else:
                                        invoiceNum = textRow
                                    foundInvoiceNum = True

                            dfFileName = pd.DataFrame({"": [fileName]})
                            dfDate = pd.DataFrame({"": [invoiceDate]})
                            dfInvoiceNum = pd.DataFrame({"": [invoiceNum]})

                            if len(pdf.pages) >= 5:
                                getTables(2,5) # checks pages 3 to 5

                            elif len(pdf.pages) == 4:
                                getTables(2,4) # checks pages 3 to 4

                            else:
                                errorMsgs.append(f"\nPdf file '{fileName}' has failed because it is less than 4 pages.")

                            if len(tables) > 1:  # if there are more than one table combines them
                                dfTables = pd.concat(tables, ignore_index=True)
                                invoiceData.append([dfFileName, dfDate, dfInvoiceNum, dfTables])

                            elif len(tables) == 1:  # if only one table just adds that
                                invoiceData.append([dfFileName, dfDate, dfInvoiceNum, tables[0]])

                            else:  # deals with no tables being found
                                # invoiceData.append([dfFileName, dfDate, dfInvoiceNum, pd.DataFrame({"": [f"Error with file table: '{fileName}'"]}), "Placeholder so that the length is 5 for error reasons"])
                                # errorMsg = createErrorMsg()
                                # errorMsg += f"\nCan't find tables"
                                # errorMsgs.append(errorMsg)

                                reader = easyocr.Reader(['en'])
                                image = pdf.pages[2].to_image(resolution=400)
                                pil_image = image.original
                                open_cv_image = np.array(pil_image)
                                result = reader.readtext(open_cv_image)

                                result.sort(key=lambda x: x[0][0][1])  # Sort the results based on the y-coordinate (helps with row detection)

                                # Group the text into rows based on y-coordinate
                                rows = []
                                current_row = []
                                previous_y = None
                                row_threshold = 10  # Threshold to consider if a new line belongs to the same row

                                for detection in result:
                                    text = detection[1]
                                    y = detection[0][0][1]  # Get the y-coordinate of the text

                                    # If the y-coordinate difference is large, it indicates a new row
                                    if previous_y is None or abs(previous_y - y) > row_threshold:
                                        if current_row:
                                            rows.append(current_row)  # Save the current row
                                        current_row = [text]  # Start a new row
                                    else:
                                        current_row.append(text)  # Add the text to the current row

                                    previous_y = y

                                if current_row:  # Append the last row after finishing the loop
                                    rows.append(current_row)

                                invoiceData.append([dfFileName, dfDate, dfInvoiceNum, pd.DataFrame(rows)])

                        except Exception as e:  # stores error info
                            dfError = pd.DataFrame({"": [f"\nError with file: '{fileName}'", e]})
                            invoiceData.append([dfError])
                            errorMsg = createErrorMsg()
                            errorMsg += f"\nPython error: '{e}'"
                            errorMsgs.append(errorMsg)


while True:  # allows the user to rewrite if they are in the file
    if len(invoiceData) > 0:  # checks there is data to write
        try:
            with pd.ExcelWriter("output.xlsx", engine="openpyxl") as writer:  # writes the data to output file
                # write headers
                pd.DataFrame({"": ["FileName"]}).to_excel(writer, sheet_name="VAT", index=False, startrow=0, startcol=0, header=False)
                pd.DataFrame({"": ["Invoice Date"]}).to_excel(writer, sheet_name="VAT", index=False, startrow=0, startcol=1, header=False)
                pd.DataFrame({"": ["Invoice number"]}).to_excel(writer, sheet_name="VAT", index=False, startrow=0, startcol=2, header=False)
                pd.DataFrame({"": ["VAT"]}).to_excel(writer, sheet_name="VAT", index=False, startrow=0, startcol=columnWithVAT+3, header=False)

                for data in invoiceData:  # displays the data
                    if len(data) == 4:
                        table = data[3].apply(lambda col: col.str.replace("\n", "", regex=False) if col.dtype == "object" else col)

                        for i in range(1, len(data[3])+1):  # put the date and invoice num on each row of the table (starts at 1 cos headers start at 0)
                            data[0].to_excel(writer, index=False, sheet_name="VAT", startrow=vatRowCount + i, startcol=0, header=False)  # filename
                            data[1].to_excel(writer, index=False, sheet_name="VAT", startrow=vatRowCount + i, startcol=1, header=False)  # date
                            data[2].to_excel(writer, index=False, sheet_name="VAT", startrow=vatRowCount + i, startcol=2, header=False)  # invoice num

                            # extracting the VAT as a separate column
                            tableTemp = table.apply(lambda col: col.str.replace(",", "", regex=False) if col.dtype == "object" else col)

                            try:
                                cell = tableTemp.iloc[i-1, columnWithVAT]

                                if cell is not None and cell.find("VAT") != -1 and cell.find(wordAfterVAT) != -1:
                                    vatCell = float(cell[cell.find("VAT")+6:cell.find(wordAfterVAT)])
                                    pd.DataFrame({"": [vatCell]}).to_excel(writer, sheet_name="VAT", index=False, startrow=vatRowCount + i, startcol=3 + len(tableTemp.columns), header=False)  # vat column

                            except Exception as e:
                                pass
                                # print(f"Failed finding VAT: {e}")

                        # table
                        table.to_excel(writer, sheet_name="VAT", index=False, startrow=vatRowCount, startcol=3)
                        vatRowCount += len(table) + blankLineCount

                for data in invoiceData:  # displays the errors at the end
                    if len(data) == 1 or len(data) == 5:
                        data[0].to_excel(writer, sheet_name="Errors", index=False, startrow=errorsRowCount, startcol=0, header=False)
                        vatRowCount += blankLineCount
            break

        except Exception as e:
            print(e)
            input("\nIf the file is open, close it and press enter")

    else:
        break

file = open("log.txt", "w")

for errorMsg in errorMsgs:
    print(errorMsg)
    file.writelines(errorMsg)

if len(errorMsgs) > 0:
    print(f"\nFailed Files: {len(errorMsgs)}")
    file.writelines(f"\nFailed Files: {len(errorMsgs)}")

file.close()