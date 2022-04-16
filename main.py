from layout import *
import companies as c
from openpyxl import load_workbook
import os

SESSIONNAME = 'Technical Day ONE: Engineering and IT - Wednesday, Feb 9, 10:00 am - 3:00 pm EST' #CHANGE THIS LINE!
LAYOUTFILENAME = 'CRC floor diagram 8x16 v6-testing.xlsx'
COMPANIESFILENAME = 'registered as of 2022.01.21.xlsx'
#'Technical Day ONE: Engineering and IT - Tuesday, Sep 14, 9:00 am - 2:00 pm EDT'
#'Technical Day TWO: Engineering and IT - Wednesday, Sep 15, 9:00 am - 2:00 pm EDT'
#'Professional Day: Business and Arts and Sciences - Monday, Sep 13, 9:00 am - 2:00 pm EDT'

if __name__ == '__main__':
    wb = load_workbook(filename = LAYOUTFILENAME)
    sheet = wb.worksheets[0]

    excludedIndustries = [] 
    exclInd = input('Enter an industry to be excluded from this layout (Enter nothing to generate layout): ')
    while exclInd != '':
        excludedIndustries = excludedIndustries + [exclInd.title()]
        exclInd = input('Enter an industry to be excluded from this layout (Enter nothing to generate layout): ')

    print('\n\nDo you want to sort by big companies? \nNOTE: This requires a column called Big Company to be manually added')
    print('with a 1 entered for any Big Companies (as determined by the Career Dev Team). All other cells in the column should be left blank.')
    SORT_BIG_COMPS = int(input('Enter 0 for no, or 1 for yes: '))

    startrow = 1
    endrow = len(list(sheet.rows)) 
    startcol = 1 
    endcol = len(list(sheet.columns)) 
    stdBooths = makeBooths(sheet, startrow, endrow, startcol, endcol)

    premBooths = filterPremiumBooths(stdBooths)

    powBooths = filterPowerBooths(stdBooths)

    stdBooths.sort(key=lambda x: x.boothName) 

    finBooths = []
    comps = c.getSortedCompanies(COMPANIESFILENAME, SESSIONNAME, excludedIndustries, SORT_BIG_COMPS)

    if len(comps) > (len(premBooths) + len(powBooths) + len(stdBooths)):
        print('WARNING: NOT ENOUGH BOOTHS IN THE LAYOUT')

    standardBoothIndex = 0
    count = 0
    unassignedComps = []

    notEnoughPrem = False
    notEnoughPower = False
    notEnoughBooths = False

    numBoothsAssigned = 0
    for comp in comps:
        if 'Premium Booth' in comp.boothType:
            print(f'{comp.boothType}: {comp.companyName}')
            try:
                pBooth = premBooths[-1]
                pBooth.companyName = comp.companyName
                finBooths = finBooths + [pBooth]
                premBooths = premBooths[0:-1]
                numBoothsAssigned += 1
            except IndexError:
                print(f'WARNING: PREMIUM BOOTH UNABLE TO BE ASSIGNED TO {comp.companyName}')
                notEnoughPrem = True
                unassignedComps = unassignedComps + [comp]

        elif comp.needsElectric:
            print(f'Needs Electric: {comp.companyName}')
            try:
                powBooth = powBooths[-1]
                powBooth.companyName = comp.companyName
                finBooths = finBooths + [powBooth]
                powBooths = powBooths[0:-1]
                numBoothsAssigned += 1
            except IndexError:
                print(f'WARNING: POWER BOOTH UNABLE TO BE ASSIGNED TO {comp.companyName}')
                notEnoughPower = True
                unassignedComps = unassignedComps + [comp]

        elif 'Standard Booth' in comp.boothType:
            print(f'{comp.boothType}: {comp.companyName}')
            # search for an empty standard booth
            try:
                stdBooth = stdBooths[-1]
                stdBooth.companyName = comp.companyName
                finBooths = finBooths + [stdBooth]
                stdBooths = stdBooths[0:-1]
                numBoothsAssigned += 1

            # don't have any standard booths left - try to find an empty power booth we can assign a standard company to
            except IndexError:
                try:
                    powBooth = powBooths[-1]
                    powBooth.companyName = comp.companyName
                    finBooth = finBooths + [powBooth]
                    powBooths = powBooths[0:-1]
                    numBoothsAssigned += 1

                # don't have any power booths left either
                except:
                    print(f'WARNING: BOOTH UNABLE TO BE ASSIGNED TO {comp.companyName}')
                    notEnoughBooths = True
                    unassignedComps = unassignedComps + [comp]

    print('--------------------------------------------------------------')
    print('FINISHED!')
    print(f'Total number of companies found: {len(comps)}')
    print(f'Total number of booths assigned: {numBoothsAssigned}\n')

    if notEnoughPrem or notEnoughPower or notEnoughBooths:
        printUnassignedComps(unassignedComps)

    # assign the booths to the excel file
    wb = load_workbook(filename = LAYOUTFILENAME)
    sheet_ranges = wb.worksheets[0]

    for b in finBooths:
        cell = sheet_ranges.cell(column=b.boothCol, row=b.boothRow)
        cell.value = b.companyName

    # booths are assigned, now figure out where to save the results

    # create a results folder (if it doesn't already exist)
    if not os.path.isdir('results'):
        os.mkdir('results')

    # create the file name
    cwd = os.getcwd()
    outputFileName = LAYOUTFILENAME[0:-5] + '_RESULTS'
    extension = '.xlsx'
    outputPath = cwd + '\\results\\' + outputFileName + extension
    
    # if we already have a results file with this name, start looking for a number
    # to stick on the end of the filename to make it unique
    # e.g. My_Booth_Layout_RESULTS (1).xlsx
    count = 1
    while os.path.exists(outputPath):
        outputPath = cwd + '\\results\\' + outputFileName + f' ({count})' + extension
        count += 1
    # found a unique name - now save it
    wb.save(outputPath)
    print(f'Result saved to: {outputPath}\n')
