import openpyxl

def createPrimaryExit(wb) -> str:
    answer = '\n'
    primarySheet = wb.get_sheet_by_name('F1 Primary List')
    for i in range(2,primarySheet.max_row+1):
        pe = primarySheet.cell(row=i, column=1).value
        if (pe != None):
            answer += 'let ' + pe.lower() + ' = PrimaryExit(strId: "' + pe + '")\n'
    return answer

def createZonePrimaries(wb) -> str:
    answer = '\n'
    zoneSheet = wb.get_sheet_by_name('F1 Zone Info')
    for i in range(2,zoneSheet.max_row + 1):
        zone = "".join(zoneSheet.cell(row=i, column=1).value.split()).lower()
        primary = zoneSheet.cell(row=i, column=2).value.strip(' ').lower()
        answer += 'let ' + zone + 'Primaries = [' + primary + ']\n'
    return answer

def createSecondaryPoints(wb) -> str:
    answer = '\n'
    d = {}
    psSheet = wb.get_sheet_by_name('F1 Primary-Secon.')
    for i in range(2,psSheet.max_row + 1):
        primary = psSheet.cell(row=i, column=1).value
        secondary = psSheet.cell(row=i, column=2).value
        edgeValue = psSheet.cell(row=i, column=3).value
        if (secondary not in d):
            d[secondary] = []
        d[secondary].append((primary,edgeValue))
    for k,v in d.items():
        answer += 'let ' + k.lower() + 'p = ['
        for primaryEdge in v:
            answer += primaryEdge[0].lower() + ' : ' + str("{0:0.1f}".format(primaryEdge[1])) + ', '
        answer = answer[:-2] + ']\n'
        
    return answer

def createSecondaryExit(wb) -> str:
    answer = '\n'
    secondarySheet = wb.get_sheet_by_name('F1 Secondary List')
    for i in range(2,secondarySheet.max_row + 1):
        secondary = secondarySheet.cell(row=i, column=1).value
        name = secondarySheet.cell(row=i, column=4).value
        answer += 'let ' + secondary.lower() + ' = SecondaryExit(id: "' + secondary + '", locationName: "' + name + '", toPrimaryMap: ' + secondary.lower() + 'p)\n'
    return answer

def createZoneSecondaries(wb) -> str:
    answer = '\n'
    zoneSheet = wb.get_sheet_by_name('F1 Zone Info')
    for i in range(2,zoneSheet.max_row + 1):
        zone = "".join(zoneSheet.cell(row=i, column=1).value.split()).lower()
        secondary = zoneSheet.cell(row=i, column=3).value.strip(' ').lower()
        answer += 'let ' + zone + 'Secondaries = [' + secondary + ']\n'
    return answer

def createZone(wb) -> str:
    answer = '\n'
    zoneSheet = wb.get_sheet_by_name('F1 Zone Info')
    for i in range(2,zoneSheet.max_row + 1):
        zone = zoneSheet.cell(row=i, column=1).value
        name = "".join(zone.split()).lower()
        answer += 'let ' + name + ' = Zone(name: "' + zone + '", primaryExits: ' + name + 'Primaries, secondaryExits: ' + name + 'Secondaries)\n'
    return answer

if __name__ == '__main__':
    wb = openpyxl.load_workbook('scesf1map.xlsx')
    pe = createPrimaryExit(wb)
    zp = createZonePrimaries(wb)
    sp = createSecondaryPoints(wb)
    se = createSecondaryExit(wb)
    zs = createZoneSecondaries(wb)
    z = createZone(wb)
    file = open('IndoorSearchSimulator.txt','w')
    file.write(pe)
    file.write(zp)
    file.write(sp)
    file.write(se)
    file.write(zs)
    file.write(z)
    file.close()
