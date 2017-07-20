import openpyxl

def zonePrimaryDirections(wb,rows) -> str:
    answer = ""
    primarySet = set()
    zpDirection = wb.get_sheet_by_name('F1 Zone-Primary Start Direction')
    startZone = ""
    name = ""
    for i in range(2,rows):
        if (zpDirection.cell(row=i, column=1).value != None):
            startZone = zpDirection.cell(row=i, column=1).value
            name = 'z' + startZone[5:]
            if ('Out' in startZone):
                name = 'oz' + startZone[8:]
            primary = zpDirection.cell(row=i, column=2).value
            d1 = zpDirection.cell(row=i, column=3).value
            d1arrow = zpDirection.cell(row=i, column=4).value
            d1image = "https://www.transparenttextures.com/patterns/asfalt-light.png"
            d2 = zpDirection.cell(row=i, column=6).value
            d2arrow = zpDirection.cell(row=i, column=7).value
            d2image = "https://www.transparenttextures.com/patterns/asfalt-light.png"
            d3 = zpDirection.cell(row=i, column=9).value
            d3arrow = zpDirection.cell(row=i, column=10).value
            d3image = "https://www.transparenttextures.com/patterns/asfalt-light.png"

            if not (startZone in primarySet):
                answer += 'var ' + name + 'Neighbors = [String:[PathStep]]()\n'
                primarySet.add(startZone)
            if (d1 != 'N/A'):
                answer += 'tempPathStepList = [PathStep]()\ntempPathStepList.append(PathStep(directionText: ' + d1 + ', directionImage: "' + d1image + '", arrow: .' + convertArrow(d1arrow) + '))\n'
            if (d2 != 'N/A'):
                answer += 'tempPathStepList.append(PathStep(directionText: ' + d2 + ', directionImage: "' + d2image + '", arrow: .' + convertArrow(d2arrow) + '))\n'
            if (d3 != 'N/A'):
                answer += 'tempPathStepList.append(PathStep(directionText: ' + d3 + ', directionImage: "' + d3image + '", arrow: .' + convertArrow(d3arrow) + '))\n'
            answer += name + 'Neighbors["' + primary + '"] = tempPathStepList\n\n'
        else:
            answer += 'primaryToPrimaryDescriptionsMap["' + startZone + '"] = ' + name + 'Neighbors\n\n\n'

    return answer


def primaryPrimaryDirections(wb,rows) -> str:
    answer = ""
    primarySet = set()
    ppDirection = wb.get_sheet_by_name('F1 Primary-Primary Directions')
    startPrimary = ""
    name = ""

    for i in range(2,rows):
        if (ppDirection.cell(row=i, column=1).value != None):
            startPrimary = ppDirection.cell(row=i, column=1).value.lower().strip()
            endPrimary = ppDirection.cell(row=i, column=2).value
            d1 = ppDirection.cell(row=i, column=3).value
            d1arrow = ppDirection.cell(row=i, column=4).value
            d1image = "https://www.transparenttextures.com/patterns/asfalt-light.png"
            d2 = ppDirection.cell(row=i, column=6).value
            d2arrow = ppDirection.cell(row=i, column=7).value
            d2image = "https://www.transparenttextures.com/patterns/asfalt-light.png"
            d3 = ppDirection.cell(row=i, column=9).value
            d3arrow = ppDirection.cell(row=i, column=10).value
            d3image = "https://www.transparenttextures.com/patterns/asfalt-light.png"
            d4 = ppDirection.cell(row=i, column=12).value
            d4arrow = ppDirection.cell(row=i, column=13).value
            d4image = "https://www.transparenttextures.com/patterns/asfalt-light.png"

            if not (startPrimary in primarySet):
                answer += 'var ' + startPrimary + 'Neighbors = [String:[PathStep]]()\n'
                primarySet.add(startPrimary)
            if (d1 != 'N/A'):
                answer += 'tempPathStepList = [PathStep]()\ntempPathStepList.append(PathStep(directionText: ' + d1 + ', directionImage: "' + d1image + '", arrow: .' + convertArrow(d1arrow) + '))\n'
            if (d2 != 'N/A'):
                answer += 'tempPathStepList.append(PathStep(directionText: ' + d2 + ', directionImage: "' + d2image + '", arrow: .' + convertArrow(d2arrow) + '))\n'
            if (d3 != 'N/A'):
                answer += 'tempPathStepList.append(PathStep(directionText: ' + d3 + ', directionImage: "' + d3image + '", arrow: .' + convertArrow(d3arrow) + '))\n'
            if (d4 != 'N/A'):
                answer += 'tempPathStepList.append(PathStep(directionText: ' + d4 + ', directionImage: "' + d4image + '", arrow: .' + convertArrow(d4arrow) + '))\n'
            answer += startPrimary + 'Neighbors["' + endPrimary + '"] = tempPathStepList\n\n'

    return answer



def primarySecondaryDirections(wb,rows) -> str:
    answer = "\n"
    primarySet = set()
    psDirection = wb.get_sheet_by_name('F1 Primary-Second. Directions')
    startPrimary = ""
    name = ""

    for i in range(2,rows):
        if (psDirection.cell(row=i, column=1).value != None):
            startPrimary = psDirection.cell(row=i, column=1).value.lower().strip()
            endPrimary = psDirection.cell(row=i, column=2).value
            d1 = psDirection.cell(row=i, column=3).value
            d1arrow = psDirection.cell(row=i, column=4).value
            d1image = "https://www.transparenttextures.com/patterns/asfalt-light.png"
            d2 = psDirection.cell(row=i, column=6).value
            d2arrow = psDirection.cell(row=i, column=7).value
            d2image = "https://www.transparenttextures.com/patterns/asfalt-light.png"
            d3 = psDirection.cell(row=i, column=9).value
            d3arrow = psDirection.cell(row=i, column=10).value
            d3image = "https://www.transparenttextures.com/patterns/asfalt-light.png"

            if (d1 != 'N/A'):
                answer += 'tempPathStepList = [PathStep]()\ntempPathStepList.append(PathStep(directionText: ' + d1 + ', directionImage: "' + d1image + '", arrow: .' + convertArrow(d1arrow) + '))\n'
            else:
                answer += 'tempPathStepList = [PathStep]()\n'
            if (d2 != 'N/A'):
                answer += 'tempPathStepList.append(PathStep(directionText: ' + d2 + ', directionImage: "' + d2image + '", arrow: .' + convertArrow(d2arrow) + '))\n'
            if (d3 != 'N/A'):
                answer += 'tempPathStepList.append(PathStep(directionText: ' + d3 + ', directionImage: "' + d3image + '", arrow: .' + convertArrow(d3arrow) + '))\n'
            answer += startPrimary + 'Neighbors["' + endPrimary + '"] = tempPathStepList\n\n'

    return answer


def addToPrimaryPrimaryMap(wb) -> str:
    answer = ""
    primarySheet = wb.get_sheet_by_name('F1 Primary List')
    for i in range(2,primarySheet.max_row):
        if (primarySheet.cell(row=i, column=1).value != None):
            answer += 'primaryToPrimaryDescriptionsMap["' + primarySheet.cell(row=i, column=1).value + '"] = ' + primarySheet.cell(row=i, column=1).value.lower() + 'Neighbors\n'
    return answer
    


def convertArrow(arrow) -> str:
    arrow = arrow.lower().strip()
    if (arrow.strip() == 'forward'):
        return 'forward'
    if (arrow == 'right'):
        return 'right'
    if (arrow == 'left'):
        return 'left'
    if (arrow == 'slight right'):
        return 'slightRight'
    if (arrow == 'slight left'):
        return 'slightLeft'
    if (arrow == 'right u-turn'):
        return 'rightUTurn'
    if (arrow == 'left u-turn'):
        return 'leftUTurn'
    return 'unknown'

if __name__ == '__main__':
    wb = openpyxl.load_workbook('scesf1map.xlsx')
    #print(wb.get_sheet_names())
    zpd = zonePrimaryDirections(wb,60)
    ppd = primaryPrimaryDirections(wb,118)
    psd = primarySecondaryDirections(wb,149)
    p = addToPrimaryPrimaryMap(wb)
    # print(zpDirection.max_row)
    # print(zpDirection.max_column)
    file = open('NavigationResult.txt','w')
    file.write('var tempPathStepList = [PathStep]()\n\n')
    file.write(zpd)
    file.write(ppd)
    file.write(psd)
    file.write(p)
    file.close()
