import xlsxwriter
import argparse

# Gather input and return nested dictionary structure
def gatherInput():
    numModules = input("How many modules are you studying this academic year? : ")
    numModules = int(numModules)

    numCredits = input("How many credits are there in total this academic year? : ")
    numCredits = int(numCredits)

    yearInfo = {'numModules': numModules, 'numCredits': numCredits, 'modules': []}

    for module in range(numModules):
        moduleName = input("What is the name of module {}? : ".format(module+1))

        moduleWeight = input("How many credits is module {} worth? : ".format(moduleName))
        moduleWeight = int(moduleWeight)/numCredits

        numModuleParts = input("How many separate marked items (Exams, Coursework, etc) are in {}? : ".format(moduleName))
        numModuleParts = int(numModuleParts)

        moduleInfo = {'name': moduleName, 'weight':moduleWeight, 'numModuleParts': numModuleParts, 'moduleParts': []}

        for item in range(numModuleParts):
            itemName = input("What is the name of module item {}? : ".format(item+1))

            weighting = input("What is the value, as a percentage of the module, of {}? : ".format(itemName))
            weighting = float(weighting)/100

            itemInfo = {'name': itemName, 'weighting': weighting}

            moduleInfo['moduleParts'].append(itemInfo)

        yearInfo['modules'].append(moduleInfo)

    return yearInfo


def generateExcel(template):
    # Create a workbook and add a worksheet
    workbook = xlsxwriter.Workbook('gradeCalculator.xlsx')
    worksheet = workbook.add_worksheet()

    fHeading = workbook.add_format({'bold': True, 'border': '0', 'bg_color': '#CAE1E8'})
    fBottom = workbook.add_format({'bold': True, 'border': '0', 'bg_color': '#8DE5FF'})
    fBottomPercent = workbook.add_format({'border': '0', 'bg_color': '#8DE5FF', 'num_format': '0.00%'})
    fData = workbook.add_format({'bg_color': '#E9E9E9'})
    fDataPercent = workbook.add_format({'bg_color': '#E9E9E9', 'num_format': '0.00%'})

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    worksheet.write(row, col, 'Num. modules:', fHeading)
    worksheet.write(row, col+1, template['numModules'], fHeading)
    row+=1
    worksheet.write(row, col, 'Num. credits:', fHeading)
    worksheet.write(row, col+1, template['numCredits'], fHeading)
    row+=2

    # For every module
    for module in range(template['numModules']):
        # Set current module using index
        currentModule = template['modules'][module]
        # Set column headings
        worksheet.write(row, col, currentModule['name'], fHeading)
        worksheet.write(row, col+1, 'Weighting', fHeading)
        worksheet.write(row, col+2, 'Grade', fHeading)
        worksheet.write(row, col+3, 'Weighted Grade', fHeading)
        row+=1
        startRange = row+1
        # For each piece of coursework, exam etc
        for modulePart in range(currentModule['numModuleParts']):
            # Set current item
            currentModulePart = currentModule['moduleParts'][modulePart]
            # Record first row
            worksheet.write(row, col, currentModulePart['name'], fData)
            worksheet.write(row, col+1, currentModulePart['weighting'], fDataPercent)
            worksheet.write(row, col+2, '', fDataPercent)
            worksheet.write(row, col+3, '=B{0}*C{0}'.format(row+1), fDataPercent)
            row+=1

        worksheet.write(row, col, 'Average grade:', fBottom)
        worksheet.write(row, col+1, '=IF(COUNTBLANK(C{0}:C{1})={2},"", AVERAGE(C{0}:C{1}))'.format(startRange,(startRange + (currentModule['numModuleParts']-1)),currentModule['numModuleParts']), fBottomPercent)
        worksheet.write(row, col+2, 'Weighted Total:', fBottom)
        worksheet.write(row, col+3, '=SUM(D{0}:D{1})'.format(startRange,(startRange + (currentModule['numModuleParts']-1))), fBottomPercent)
        row+=1
        worksheet.write(row, col+2, 'Module-Weighted total:', fBottom)
        worksheet.write(row, col+3, '={0}*D{1}'.format(currentModule['weight'], row), fBottomPercent)
        row+=2

    workbook.close()


# Iterate over the data and write it out row by row.
generateExcel(gatherInput())
