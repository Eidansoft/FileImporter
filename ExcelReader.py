import xlrd
import warnings

class Field:
    def __init__(self, row, col):
        self.row = row
        self.col = col

    def __repr__(self):
        return "({}, {})".format(self.row, self.col)


class Reader:

    def __init__(self, path, header_field_up_left, header_field_down_rigth, sheet=0):

        # import ipdb; ipdb.set_trace(context=21)

        if (header_field_up_left.row != header_field_down_rigth.row and
            header_field_up_left.col != header_field_down_rigth.col):
            raise ValueError("The headers fields must be in the same row or in the same column!")
        elif (header_field_up_left.col > header_field_down_rigth.col or
              header_field_up_left.row > header_field_down_rigth.row):
            raise ValueError("The first header field must be the top-left one!")

        self.horizontal = False
        if header_field_up_left.row == header_field_down_rigth.row:
            self.horizontal = True

        self.path = path
        self.header_field_up_left = header_field_up_left
        self.header_field_down_rigth = header_field_down_rigth
        self.sheet = sheet

        # Currently the xlrd library has some warnings about deprecation methods that I can ignore
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", category=PendingDeprecationWarning)
            warnings.filterwarnings("ignore", category=DeprecationWarning)
            wb = xlrd.open_workbook(self.path)

        # Configure the sheet to work with
        ws = wb.sheet_by_index(self.sheet)

        self.headers = []
        if self.horizontal:
            for col in range(self.header_field_up_left.col, self.header_field_down_rigth.col + 1):
                self.headers.append( ws.cell(self.header_field_up_left.row, col).value )
        else:
            for row in range(self.header_field_up_left.row, self.header_field_down_rigth.row + 1):
                self.headers.append( ws.cell(row, self.header_field_up_left.row).value )

    def get_headers(self):
        return self.headers

    def get_data(self):
        raise NotImplementedError('[ERROR] Method still not implemented.')


class DictConverter(Reader):
    def get_data(self):
        # Currently the xlrd library has some warnings about deprecation methods that I can ignore
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", category=PendingDeprecationWarning)
            warnings.filterwarnings("ignore", category=DeprecationWarning)
            wb = xlrd.open_workbook(self.path)

        ws = wb.sheet_by_index(self.sheet)

        res = []

        if self.horizontal:
            row = self.header_field_up_left.row + 1
            while row < ws.nrows:
                obj = dict()
                for col in range(self.header_field_up_left.col, self.header_field_down_rigth.col + 1):
                    obj[ ws.cell(self.header_field_up_left.row, col).value ] = ws.cell(row, col)

                res.append(obj)
                row += 1
        else:
            col = self.header_field_up_left.col + 1
            while col < ws.ncols:
                obj = dict()
                for row in range(self.header_field_up_left.row, self.header_field_down_rigth.row + 1):
                    obj[ ws.cell(row, self.header_field_up_left.col).value ] = ws.cell(row, col)

                res.append(obj)
                col += 1

        # Clean the values to contain the proper value instead a Cell object
        # more info at https://xlrd.readthedocs.io/en/latest/api.html#xlrd.sheet.Cell
        for obj in res:
            for element in self.get_headers():
                if (obj[element].ctype == 0 or obj[element].ctype == 6):
                    # Empty string or empty cell
                    obj[element] = None
                elif (obj[element].ctype == 1 or obj[element].ctype == 2):
                    # String or number
                    obj[element] = obj[element].value
                elif obj[element].ctype == 4:
                    # Boolean
                    if obj[element].value == 1:
                        obj[element] = True
                    else:
                        obj[element] = False
                elif obj[element].ctype == 3:
                    # Date, more info: https://xlrd.readthedocs.io/en/latest/api.html#xlrd.xldate.xldate_as_datetime
                    obj[element] = xlrd.xldate.xldate_as_datetime(obj[element].value, 0)
                elif obj[element].ctype == 5:
                    # Excel error
                    obj[element] = xlrd.biffh.error_text_from_code[obj[element].value]

        return res
