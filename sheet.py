import string
class Cell:
    EMPTY_CELL = '_'
    EMPTY = ''

    def __init__(self, value=None, formula=None):
        self.value = value
        self.formula = formula

    def get_value(self):
        return self.value

    def __str__(self):
        return str(self.value) if self.value is not None else self.EMPTY_CELL

    def cell_loc(self, name):
        """gets coordinate for a cell as text and returns a tuple of the cell location"""
        try:
            first_letter = name[0].upper()
            if not ('A' <= first_letter <= 'Z'):
                raise ValueError("Error: First part of the cell name must be a letter.")

            column_guess_num = ord(first_letter) - ord('A') + 1
            row_guess = int(name[1:])
            user_guess = (row_guess - 1, column_guess_num - 1)
            return user_guess

        except ValueError:
            print("Error: Second part of the cell name must be a number.")
            return False

class Sheet:
    def __init__(self, rows, cols):
        self.rows = rows
        self.cols = cols
        self.cells_dict = {}
        self.formula_cells = {}  # Dictionary to store cells with formulas

        # Initialize cells_dict with empty values for all cells
        for row in range(rows):
            for col in range(cols):
                cell_address = self.get_cell_address(row, col)
                self.cells_dict[cell_address] = {'value': '', 'formula': None}

    def __str__(self):
        result = ""
        # Generating column labels (A, B, C, ...)
        column_labels = [string.ascii_uppercase[i] for i in range(self.cols)]

        # Printing column labels as the first row
        result += "   "  # Padding for row index column
        for label in column_labels:
            result += str(label).ljust(8)  # Adjust width to align cells properly
        result += "\n"  # Start a new line for the next row

        # Printing the cells
        for row_index in range(self.rows):
            result += str(row_index + 1).ljust(3)  # Print row index at the beginning of each row
            for col_index in range(self.cols):
                cell_address = self.get_cell_address(row_index, col_index)
                value = self.cells_dict.get(cell_address, {}).get('value', Cell.EMPTY_CELL)
                result += str(value).ljust(8)  # Adjust width to align cells properly
            result += "\n"

        return result.rstrip("\n")  # Remove trailing newline before returning

    def get_cell_address(self, row, col):
        return f"{string.ascii_uppercase[col]}{row + 1}"

    def set_value(self, value, row, col):
        """insert a value to a cell"""
        if row < 0 or row >= self.rows or col < 0 or col >= self.cols:
            raise IndexError("Row or column index is out of range")

        cell_address = self.get_cell_address(row, col)
        self.cells_dict[cell_address]['value'] = value
        # Only update formula if the cell doesn't already have a formula
        if self.cells_dict[cell_address]['formula'] is None:
            self.cells_dict[cell_address]['formula'] = str(value)

    def get_value(self, row, col):
        """return the stored value of a cell"""
        if row < 0 or row >= self.rows or col < 0 or col >= self.cols:
            raise IndexError("Error! Row or column index is out of range")

        cell_address = self.get_cell_address(row, col)
        value = self.cells_dict.get(cell_address, {}).get('value', Cell.EMPTY_CELL)
        if isinstance(value, str) and value.replace('.', '', 1).isdigit():
            try:
                value = float(value)
            except ValueError:
                value = int(value)
        elif value == Cell.EMPTY_CELL:
            raise ValueError("Error! Cell is empty")

        return value

    def set_formula(self, formula, row, col):
        """set the formula of a cell"""
        if row < 0 or row >= self.rows or col < 0 or col >= self.cols:
            raise IndexError("Row or column index is out of range")

        cell_address = self.get_cell_address(row, col)
        self.cells_dict[cell_address] = {'value': None, 'formula': formula}
        self.formula_cells[cell_address] = formula

    def get_formula(self, row, col):
        """return the formula of a cell"""
        try:
            if row < 0 or row >= self.rows or col < 0 or col >= self.cols:
                raise IndexError("Error! Row or column index is out of range")
            cell_address = self.get_cell_address(row, col)
            return self.cells_dict[cell_address]['formula']
        except IndexError as e:
            print(e)  # Print an error message
            return

    def get_formula_cells(self):
        return self.formula_cells

    def insert_row(self):
        """insert a row to the sheet"""
        if self.rows == self.get_max_rows():
            raise IndexError("Sheet is already full. Cannot insert a new row.")
        # Insert empty cells in the new row
        new_row_index = self.rows
        for col in range(self.cols):
            cell_address = self.get_cell_address(new_row_index, col)
            self.cells_dict[cell_address] = {'value': '', 'formula': None}
        self.rows += 1

    def get_max_rows(self):
        """Returns the maximum number of rows supported by the sheet."""
        return 50  # Example limit

    def print_sheet(self):
        """Print the sheet with values and formulas."""
        for row_index in range(self.rows):
            for col_index in range(self.cols):
                value = self.get_value(row_index, col_index)
                formula = self.get_formula(row_index, col_index)
                print(f"Cell [{row_index}, {col_index}]: Value = {value}, Formula = {formula}")