import pdfplumber
import pandas as pd
# openpyxl
import os
from contextlib import redirect_stderr

def createErrorMsg():
    Msg = f"FileName: {fileName} has failed."

    if not foundDate:
        Msg += f"\n{invoiceDate}."

    if not foundInvoiceNum:
        Msg += f"\n{invoiceNum}."

    if not foundDate or not foundInvoiceNum:
        Msg += f"\nFirstPageText\n{firstPageText}"

    return Msg


# CHANGE
blankLineCount = 1  # amount of blank lines between each pdf
firstTablePage = 3  # range of pages to load the table from
lastTablePage = 4   # range of pages to load the table from
wordAfterVAT = "Test"

# DON'T CHANGE
invoiceData = []; errorMsgs = []
fileCount = 1; rowCount = 0; colCount = 0

if not os.path.exists(os.getcwd() + "\\pdfs"):  # checks the folder pdfs exists
    os.mkdir("pdfs")

files = os.listdir(os.getcwd() + "\\pdfs")

if len(files) == 0:  # handle no pdfs
    print("No pdfs found")
    quit()

with open(os.devnull, "w") as fNull:  # remove error msgs
    with redirect_stderr(fNull):  # remove pdfplumber error msgs
        for fileName in files:
            print(f"File {fileCount}/{len(files)}")  # progress update
            fileCount+=1

            invoiceDate = f"Can't find invoice date in '{fileName}'"; invoiceNum = f"Can't find invoice no in '{fileName}'"
            tables = []
            foundDate = False; foundInvoiceNum = False

            if ".pdf" in fileName:
                    with pdfplumber.open("pdfs\\" + fileName) as pdf:
                        if len(pdf.pages) >= lastTablePage:
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

                                for pageNum in range(firstTablePage-1, lastTablePage):  # checks the selected pages
                                    table = pdf.pages[pageNum].extract_table()

                                    if table is not None:
                                        if len(table) > 0:
                                            tables.append(pd.DataFrame(table[1:], columns=table[0]))

                                if len(tables) > 1:  # if there are more than one table combines them
                                    dfTables = pd.concat(tables, ignore_index=True)
                                    invoiceData.append([dfFileName, dfDate, dfInvoiceNum, dfTables])

                                elif len(tables) == 1:  # if only one table just adds that
                                    invoiceData.append([dfFileName, dfDate, dfInvoiceNum, tables[0]])

                                else:  # deals with no tables being found
                                    invoiceData.append([dfFileName, dfDate, dfInvoiceNum, pd.DataFrame({"": [f"Error with file table: '{fileName}'"]})])
                                    errorMsg = createErrorMsg()
                                    errorMsg += f"\nCan't find tables\nPage 3: {pdf.pages[2].extract_text().split("\n")}\nPage 4: {pdf.pages[3].extract_text().split("\n")}\nPage 5: {pdf.pages[4].extract_text().split("\n")}"
                                    errorMsgs.append(errorMsg)

                            except Exception as e:  # stores error info
                                dfError = pd.DataFrame({"": [f"Error with file: '{fileName}'", e]})
                                invoiceData.append([dfError])
                                errorMsg = createErrorMsg()
                                errorMsg += f"\nPython error: '{e}'"
                                errorMsgs.append(errorMsg)

                        else:
                            errorMsgs.append(f"Pdf file '{fileName}' has failed because it is too short and cannot read all table pages.")

while True:  # allows the user to rewrite if they are in the file
    if len(invoiceData) > 0:  # checks there is data to write
        try:
            with pd.ExcelWriter("output.xlsx", engine="openpyxl") as writer:  # writes the data to output file
                for data in invoiceData:  # displays the data
                    if len(data) == 4:
                        for i in range(1, len(data[3])+1):  # put the date and invoice num on each row of the table (starts at 1 cos headers start at 0)
                            data[0].to_excel(writer, index=False, startrow=rowCount+i, startcol=0, header=False)  # filename
                            data[1].to_excel(writer, index=False, startrow=rowCount+i, startcol=1, header=False)  # date
                            data[2].to_excel(writer, index=False, startrow=rowCount+i, startcol=2, header=False)  # invoice num

                            cell = data[3].iloc[rowCount+i-1,len(data[3].columns)-1].replace("\n", "")
                            vatCell = float(cell[cell.find("VAT")+5:cell.find(wordAfterVAT)])
                            pd.DataFrame({"": [vatCell]}).to_excel(writer, index=False, startrow=rowCount+i, startcol=3 + len(data[3].columns), header=False)  # vat column

                        # write headers
                        pd.DataFrame({"": ["FileName"]}).to_excel(writer, index=False, startrow=rowCount, startcol=0, header=False)
                        pd.DataFrame({"": ["Invoice Date"]}).to_excel(writer, index=False, startrow=rowCount, startcol=1, header=False)
                        pd.DataFrame({"": ["Invoice number"]}).to_excel(writer, index=False, startrow=rowCount, startcol=2, header=False)
                        pd.DataFrame({"": ["VAT"]}).to_excel(writer, index=False, startrow=rowCount, startcol=3 + len(data[3].columns), header=False)

                        # table
                        data[3] = data[3].apply(lambda col: col.str.replace('\n', '', regex=False) if col.dtype == 'object' else col)  # removes new lines
                        data[3].to_excel(writer, index=False, startrow=rowCount, startcol=3)
                        rowCount +=len(data[3]) + blankLineCount

                for data in invoiceData:  # displays the errors at the end
                    if len(data) == 1:
                        data[0].to_excel(writer, index=False, startrow=rowCount, startcol=0, header=False)
                        rowCount += blankLineCount
            break

        except Exception as e:
            print(e)
            input("If the file is open, close it and press enter\n")

    else:
        break

for errorMsg in errorMsgs:
    print(errorMsg)

if len(errorMsgs) > 0:
    print(f"\nFailed Files: {len(errorMsgs)}")