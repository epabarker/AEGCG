import xlsxwriter

# CLASSES #####################################################################


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


class Style:
    """ Store XLS styles """

    styles = {}

    def add(self, key, value):
        """
        add new style
        :param key: style reference label
        :param value: style definition
        """
        self.styles[key] = value

    def get(self, style):
        """
        :param style: style definition to fetch
        :return style definition
        """
        return self.styles[style]


# FUNCTIONS ###################################################################


def obtain_major_details():
    """
    Obtain all details about the course one year
    :return Major object
    """

    major = Major()

    question = "How many modules are you studying this academic year? "
    total_modules = input(question)
    major.total_modules = int(total_modules)

    question = "How many credits are there in total this academic year? "
    credits = input(question)
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

        question = "What is the name of module {}? : "
        module.name = input(question.format(module_idx+1))

        question = "How many credits is module {} worth? : "
        module_weight = input(question.format(module.name))
        module.weighting = int(module_weight) / major.credits

        question = "How many assesments (Exams, Coursework, etc) are in {}? : "
        total_module_assesments = input(question.format(module.name))
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

        question = "What is the type of the assesment for the assesment #{}? "
        print(question.format(assesment_idx + 1))

        assesment_type_idx = 0
        for assesment_type in assesment.types:
            assesment_type_idx += 1
            print('{} - {}'.format(assesment_type_idx, assesment_type))
        assesment.type = assesment.types[int(input('Select [1-4]: '))-1]

        question = "What is the value, as a percentage of the module, of {}? "
        weighting = input(question.format(assesment.type))
        assesment.weighting = float(weighting)/100

        module.assesments.append(assesment)


def create_styles(workbook):
    """
    :param workbook: xlsxwriter Workbook object
    return Style object
    """

    style = Style()

    styles = {

        'heading': {
            'bold': True, 'border': '0', 'bg_color': '#CAE1E8'
        },

        'bottom': {
            'bold': True, 'border': '0', 'bg_color': '#8DE5FF'
        },

        'bottom_percent': {
            'border': '0', 'bg_color': '#8DE5FF', 'num_format': '0.00%'
        },

        'data': {
            'bg_color': '#E9E9E9', 'num_format': '0'
        },

        'data_percent': {
            'bg_color': '#E9E9E9', 'num_format': '0.00%'
        }

    }

    for key, value in styles.items():
        style.add(key, workbook.add_format(value))

    return style


def generate_spreadsheet(major):

    # Create a workbook and add a worksheet
    workbook = xlsxwriter.Workbook('grade_calculator.xlsx')
    worksheet = workbook.add_worksheet()

    xls_style = create_styles(workbook)

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    worksheet.write(row, col, 'Num. modules:', xls_style.get('heading'))
    worksheet.write(row, col+1, major.total_modules, xls_style.get('heading'))
    row += 1
    worksheet.write(row, col, 'Num. credits:', xls_style.get('heading'))
    worksheet.write(row, col+1, major.credits, xls_style.get('heading'))
    row += 2

    for module in major.modules:

        # Set column headings
        worksheet.write(row, col, module.name, xls_style.get('heading'))
        worksheet.write(row, col+1, 'Weighting', xls_style.get('heading'))
        worksheet.write(row, col+2, 'Grade', xls_style.get('heading'))
        worksheet.write(row, col+3, 'Weighted Grade', xls_style.get('heading'))

        row += 1
        start_range = row+1

        for assesment in module.assesments:

            # Record first row
            worksheet.write(row, col, assesment.type, xls_style.get('data'))
            worksheet.write(row, col+1, assesment.weighting, xls_style.get('data_percent'))
            worksheet.write(row, col+2, '', xls_style.get('data'))
            worksheet.write(row, col+3, '=B{0}*C{0}'.format(row+1), xls_style.get('data'))
            row += 1

        worksheet.write(row, col, 'Average grade:', xls_style.get('bottom'))

        assesments = len(module.assesments)

        template = '=IF(COUNTBLANK(C{0}:C{1})={2},"", AVERAGE(C{0}:C{1}))'
        end_range = (start_range + (assesments-1))
        cell_data = template.format(start_range, end_range, assesments)
        worksheet.write(row, col+1, cell_data, xls_style.get('data'))

        worksheet.write(row, col+2, 'Weighted Total:', xls_style.get('bottom'))

        template = '=SUM(D{0}:D{1})'
        cell_data = template.format(start_range, end_range)
        worksheet.write(row, col+3, cell_data, xls_style.get('data'))

        row += 1
        worksheet.write(row, col+2, 'Module-Weighted total:', xls_style.get('bottom'))

        template = '={0}*D{1}'
        cell_data = template.format(module.weighting, row)
        worksheet.write(row, col+3, cell_data, xls_style.get('data'))

        row += 2

    workbook.close()


# MAIN ########################################################################


if __name__ == '__main__':

    major = obtain_major_details()
    print('Data collection complete')

    generate_spreadsheet(major)
    print('Spreadsheet generation complete')
