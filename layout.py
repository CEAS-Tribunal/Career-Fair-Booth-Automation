class Booth(object):
    def __init__(self, boothName, boothRow, boothCol, previousBooth=None, companyName="", isPower=False, isPremium = False):
        self.boothName = boothName
        self.companyName = companyName 
        self.boothRow = boothRow
        self.boothCol = boothCol
        self.previousBooth= previousBooth
        self.isPower = isPower
        self.isPremium = isPremium
        if(self.isPremium):
            self.isPower = True

    def makePremium(self):
        self.isPremium = True
        self.isPower = True

    def show(self):
        print()
        print(str(self.boothName))
        print("Row :"+ str(self.boothRow))
        print("Col :" + str(self.boothCol))
        print("Is Power :" + str(self.isPower))
        print("Company :" + self.companyName)
            

def makeBooths(sheet, startrow, endrow, startcol, endcol):
    boothArr = [] # stores booth objects
    boothNames = [] # stores just the booth names (to make sure we don't store the same one twice)

    for col in range(startcol, endcol + 1):
        for row in range(startrow, endrow + 1):
            cell = sheet.cell(column=col, row=row)

            if (str(cell.value)) != 'None' and len(str(cell.value)) > 1: # found empty cell or column/row header
                # check if cell is truly a booth or something else
                if str(cell.value).islower(): 
                    continue # found a cell labeled 'exit' or 'check-in', things like that - skip it
                if not ((str(cell.value))[0].isalpha() and (str(cell.value))[1].isnumeric()):
                    continue # first character of a booth much be a letter, and second character must be a number - if not, skip it
                
                # found an actual booth
                booth = Booth(cell.value, row, col)

                # check if booth is premium
                if '*' in booth.boothName:
                    booth.makePremium() 

                # check if booth needs power 
                # (if it's premium, it's automatically given power. in this case, we can skip giving it power again)
                elif cell.fill.start_color.rgb == 'FFFFFF00':
                    booth.isPower = True

                # double check that we haven't (somehow) already stored this booth
                if cell.value not in boothNames:
                    boothNames = boothNames + [cell.value]
                    boothArr = boothArr + [booth]

    return boothArr

def filterPremiumBooths(boothAr):
    premBooths = []
    for index,booth in enumerate(boothAr):
        if '*' in booth.boothName:
            premBooths = premBooths + [booth]
            boothAr.pop(index)
    return premBooths

def filterPowerBooths(boothAr):
    powBooths = []
    for index, booth in enumerate(boothAr):
        if booth.isPower and (not booth.isPremium):
            powBooths = powBooths + [booth]
            boothAr.pop(index)
    return powBooths

def printUnassignedComps(unassignedComps):
    premComps = [comp for comp in unassignedComps if comp.boothType == 'Premium Booth']
    powerComps = [comp for comp in unassignedComps if comp.boothType == 'Standard Booth' and comp.needsElectric]
    stdComps = [comp for comp in unassignedComps if comp.boothType == 'Standard Booth' and not comp.needsElectric]
    if len(premComps) > 0:
        print('--------------------------------------------------------------')
        print('WARNING: NOT ENOUGH PREMIUM BOOTHS! UNASSIGNED PREMIUM COMPANIES: ')
        for premComp in premComps:
            print(premComp.companyName)
        print('--------------------------------------------------------------')
    
    if len(powerComps) > 0:
        print('--------------------------------------------------------------')
        print('WARNING: NOT ENOUGH POWER BOOTHS! UNASSIGNED POWER COMPANIES: ')
        for powerComp in powerComps:
            print(powerComp.companyName)
        print('--------------------------------------------------------------')

    if len(stdComps) > 0:
        print('--------------------------------------------------------------')
        print('WARNING: NOT ENOUGH STANDARD BOOTHS! UNASSIGNED STANDARD COMPANIES: ')
        for stdComp in stdComps:
            print(stdComp.companyName)
        print('--------------------------------------------------------------')
    

