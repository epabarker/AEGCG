import xlsxwriter

# CLASSES ######################################################################

class Major:
    """ The course the student is studying """

    total_modules = 0

    credits = 0

    modules = []


class Module:
    """ Module under the course """

    name = None

    weighting = 0

    total_assements = 0

    assesments = []


class Assesment:
    """ Module assesment details """

    types = ['Exam', 'Coursework', 'Class Test', 'Lab Test']

    type = None  # Exam, Coursework

    weighting = 0


# FUNCTIONS ####################################################################

# Gather input and return nested dictionary structure
def obtain_major_details():
    """
    Obtain all details about the course one year
    :return Major object
    """

    major = Major()

    total_modules = input("How many modules are you studying this academic year? : ")
    major.total_modules = int(total_modules)

    credits = input("How many credits are there in total this academic year? : ")
    major.credits = int(credits)

    obtain_module_details(major)

    return major


def obtain_module_details(major):
    """
    Obtain all information about major modules
    :param major: Major object
    :return void
    """

    for module_idx in range(major.total_modules):

        module = Module()

        module.name = input("What is the name of module {}? : ".format(module_idx+1))

        module_weight = input("How many credits is module {} worth? : ".format(module.name))
        module.weighting = int(module_weight) / major.credits

        total_module_assesments = input("How many assesments (Exams, Coursework, etc) are in {}? : ".format(module.name))
        module.total_assements = int(total_module_assesments)

        obtain_assesment_details(module)

        major.modules.append(module)


def obtain_assesment_details(module):
    """
    Obtain all information about module assesments
    :param module: Module object
    :return void
    """

    for assesment_idx in range(module.total_assements):

        assesment = Assesment()
        print("What is the type of the assesment for the assesment #{}? : ".format(assesment_idx + 1))

        assesment_type_idx = 0
        for assesment_type in assesment.types:
            assesment_type_idx += 1
            print('{} - {}'.format(assesment_type_idx, assesment_type))
        assesment.type = assesment.types[int(input('Select [1-4]: '))-1]

        weighting = input("What is the value, as a percentage of the module, of {}? : ".format(assesment.type))
        assesment.weighting = float(weighting)/100

        module.assesments.append(assesment)


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


# MAIN #########################################################################

if __name__ == '__main__':
    # Iterate over the data and write it out row by row.
    major = obtain_major_details()
    print(major.modules[0].name)
    print(major.modules[0].assesments[0].type)
