import openpyxl

def converter(inputFile, outputFile, startTime, date):
    actArr = []
    actArr.append([])
    animDict = {}
    rdWB = openpyxl.load_workbook(inputFile)
    book = openpyxl.load_workbook(outputFile)
    newSheet = book.worksheets[0]
    worksheet = rdWB.worksheets[0]
    """hFont = xlwt.Font()
    hFont.bold = True
    bAlign = xlwt.Alignment()
    bAlign.horz = xlwt.Alignment.HORZ_LEFT
    rAlign = xlwt.Alignment()
    rAlign.horz = xlwt.Alignment.HORZ_RIGHT
    hrStyle = xlwt.XFStyle()
    hrStyle.font = hFont
    hrStyle.alignment = bAlign
    dayStyle = xlwt.XFStyle()
    dayStyle.alignment = bAlign
    minStyle = xlwt.XFStyle()
    minStyle.alignment = rAlign
    workbook = xlwt.Workbook()
    ws = workbook.add_sheet("checkTime")
    hourNum = 0
    dayNum = 0
    while hourNum < 2400:
        minuteNum = 1
        ws.write(colNum, 1, "Day" + str(dayNum), dayStyle)
        ws.write(colNum, 2, str(hourNum), hrStyle)
        hourNum += 100
        colNum += 1
        while minuteNum < 60:
            ws.write(colNum, 1, "Day" + str(dayNum), dayStyle)
            ws.write(colNum, 2, str(minuteNum), minStyle)
            minuteNum += 1
            colNum += 1
    """
    x = 4
    animNum = worksheet.cell(row = 5, column = 3).value
    animDict[animNum] = []
    cellVal = worksheet.cell(row=(x + 1), column=1).value
    while worksheet.cell(row=(x + 1), column=1).value == "Result 1":
        if animNum == worksheet.cell(row = x + 1,column = 3).value:
            animDict[animNum].append(worksheet.cell(row=x + 1, column=6).value)
        else:
            animNum = worksheet.cell(row = x + 1, column = 3).value
            animDict[animNum] = []
            animDict[animNum].append(worksheet.cell(row=x + 1,column=6).value)
        x += 1
    animNumSheet = 2
    startI = int(startTime[:2]) * 60 + int(startTime[3:]) + 300
    animList = animDict.keys()
    animListNum = 0
    currCellVal = newSheet.cell(row=startI + 1, column=animNumSheet + 1).value
    while isinstance(currCellVal, float) or isinstance(currCellVal, int):
        currCellVal = newSheet.cell(row=startI + 1,column=animNumSheet + 1).value
        startI += 1
    currCount = 0
    for lst in animList:
        i = startI
        for count in range(len(animDict) + 1):
            if newSheet.cell(row=1, column=count + 1).value == lst:
                currCount = count
        for data in animDict[lst]:
            newSheet.cell(row=i + 1, column=currCount + 1).value= data
            currCount += 1
    rdWB.save(inputFile)
    book.save(outputFile)


def test():
    print("Test!")
    converter("ex.xlsx", "Checktime.xlsx", "19:00", "05/18/2018")


def main():
    inputFile = input("Enter input file (ex: rawInput): ")
    outputFile = input("Enter output file (ex: output): ")
    startTime = input("Please enter start time of first day (ex: 09:16,23:14): ")
    date = input("Please enter date (ex: 01/02/1970) :")
    # longInput = input("Enter daily stop times and gap between days (ex. 09:16,04;08:15,05... hh:mm,mm;hh:mm,mm): ")
    # dayArray = longInput.split(";")
    # camNum = input("Enter Camera #(ex: 1): ")
    converter(inputFile + ".xlsx", outputFile + ".xlsx", startTime, date)  # , dayArray, camNum)


test()
