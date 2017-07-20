import openpyxl

def primaryPrimaryDistance(wb) -> str:
    answer = ""
    ppSheet = wb.get_sheet_by_name('F1 Primary-Primary')
    for i in range(2,ppSheet.max_row+1):
        answer += 'graph.addEdge(from: "' + ppSheet.cell(row=i, column=1).value + '", to: "' + ppSheet.cell(row=i, column=2).value + '", weight: ' + str("{0:0.1f}".format(ppSheet.cell(row=i, column=3).value)) + ')\n'
    return answer

if __name__ == '__main__':
    wb = openpyxl.load_workbook('scesf1map.xlsx')
    pp = primaryPrimaryDistance(wb)
    file = open('Searcher.txt','w')
    file.write(pp)
    file.close()
