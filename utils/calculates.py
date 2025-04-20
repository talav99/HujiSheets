from tkinter import messagebox
from sheet import Cell
import re

class FormulaParser:
    @staticmethod
    def evaluate_formula(formula, sheet):
        """evaluating mathematical formulas"""
        formula = formula.lstrip('=').strip()

        # Split formula into function name and arguments
        parts = formula.split("(")
        if len(parts) < 2:
            messagebox.showerror("Error", "Invalid formula format!")
            return None

        function_name = parts[0].upper()
        arguments = parts[1].rstrip(")")  # Extract arguments (remove closing parenthesis)

        # Call the appropriate method from FunctionLibrary based on function name
        if function_name == "SUM":
            return FunctionLibrary.calc_sum(arguments, sheet)
        elif function_name == "AVERAGE":
            return FunctionLibrary.average(arguments, sheet)
        elif function_name == "MAX":
            return FunctionLibrary.find_max(arguments, sheet)
        elif function_name == "MIN":
            return FunctionLibrary.find_min(arguments, sheet)
        else:
            # Handle unsupported function
            messagebox.showerror("Error", "Unsupported function: " + function_name)
            return None

    @staticmethod
    def convert_to_math_formula(formula, sheet):
        """Split the formula into its components"""
        parts = formula.split()
        print("parts,", parts)
        if len(parts) == 1:
            cell_name1 = parts[0]  # copy the value from the cell
            coeff1, cell_name1 = FormulaParser.extract_coefficient(cell_name1)
            result = coeff1 * FormulaParser.evaluate_cell_reference(cell_name1, sheet)
            return result
        elif len(parts) != 3:
            raise ValueError("Invalid formula format! Formula should be in the form '<cell1> <operation> <cell2>"
                             "for example, '=A2 + B4'")

        cell1, operation, cell2 = parts

        # Extract coefficients from cell names
        coeff1, cell_name1 = FormulaParser.extract_coefficient(cell1)
        coeff2, cell_name2 = FormulaParser.extract_coefficient(cell2)

        # Process cell references to get their values
        value1 = FormulaParser.evaluate_cell_reference(cell_name1, sheet)
        value2 = FormulaParser.evaluate_cell_reference(cell_name2, sheet)

        # Apply coefficients after checking for None
        if value1 is not None:
            value1 *= coeff1
        if value2 is not None:
            value2 *= coeff2

        try:
            # Perform the mathematical operation based on the provided operation
            if operation == '+':
                result = value1 + value2
            elif operation == '-':
                result = value1 - value2
            elif operation == '*':
                result = value1 * value2
            elif operation == '/':
                if value2 == 0:
                    raise ZeroDivisionError  # Raise ZeroDivisionError explicitly
                result = value1 / value2
            else:
                raise ValueError("Invalid mathematical operation specified!")

            return result
        except ZeroDivisionError:  # Catch ZeroDivisionError
            raise ValueError("Division by zero error!")

    @staticmethod
    def evaluate_cell_reference(cell_name, sheet):
        if cell_name.startswith('='):
            return FormulaParser.convert_to_math_formula(cell_name[1:], sheet)
        else:
            # Otherwise, directly retrieve the value from the sheet
            row, col = Cell().cell_loc(cell_name)
            if row < 0 or row >= sheet.rows or col < 0 or col >= sheet.cols:
                raise ValueError(f"Cell '{cell_name}' is out of sheet dimensions.")
            value = sheet.get_value(row, col)
            # Return 0 if the cell has no value or contains a string
            return float(value) if isinstance(value, (int, float)) else 0

    @staticmethod
    def extract_coefficient(cell):
        """extracts the coefficient (if any) from a cell reference."""
        coeff = 1
        cell_name = cell
        if cell[0].isdigit():
            coeff_str = ''
            for char in cell:
                if char.isdigit():
                    coeff_str += char
                else:
                    break
            coeff = int(coeff_str)
            cell_name = cell[len(coeff_str):]  # Remove coefficient from cell name
        return coeff, cell_name

class FunctionLibrary:
    @staticmethod
    def is_valid_range(cell_range, sheet):
        """Check if the given cell range is valid within the sheet dimensions."""
        start, end = cell_range.split(":")
        start_row, start_col = Cell().cell_loc(start)
        end_row, end_col = Cell().cell_loc(end)

        if (start_row < 0 or start_row >= sheet.rows or start_col < 0 or start_col >= sheet.cols
                or end_row < 0 or end_row >= sheet.rows or end_col < 0 or end_col >= sheet.cols):
            return False
        return True

    @staticmethod
    def calc_sum(cell_range, sheet):
        """Calculate the sum of values within the given cell range."""
        if not FunctionLibrary.is_valid_range(cell_range, sheet):
            raise ValueError("Invalid cell range.")
        start, end = cell_range.split(":")
        start_row, start_col = Cell().cell_loc(start)
        end_row, end_col = Cell().cell_loc(end)
        total = 0
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                current_val = sheet.get_value(row, col)
                if current_val is not None and (type(current_val) == int or type(current_val) == float):
                    total += current_val
        return total

    @staticmethod
    def find_cell_len(cell_range):
        """Calculates how many cells are in a range"""
        start, end = cell_range.split(":")
        start_row, start_col = Cell().cell_loc(start)
        end_row, end_col = Cell().cell_loc(end)
        return (end_row - start_row + 1) * (end_col - start_col + 1)

        # Calculate the total number of cells in the range
        total = num_rows * num_cols
        return total

    @staticmethod
    def average(cell_range, sheet):
        """Calculate the average of values within the given cell range."""
        if not FunctionLibrary.is_valid_range(cell_range, sheet):
            raise ValueError("Invalid cell range.")
        cell_sum = FunctionLibrary.calc_sum(cell_range, sheet)
        num_of_cells = FunctionLibrary.find_cell_len(cell_range)
        return cell_sum / num_of_cells

    @staticmethod
    def find_max(cell_range, sheet):
        """Find the maximum value within the given cell range."""
        if not FunctionLibrary.is_valid_range(cell_range, sheet):
            raise ValueError("Invalid cell range.")
        start, end = cell_range.split(":")
        start_row, start_col = Cell().cell_loc(start)
        end_row, end_col = Cell().cell_loc(end)

        max_value = max(
            sheet.get_value(row, col)
            for row in range(start_row, end_row + 1)
            for col in range(start_col, end_col + 1)
            if isinstance(sheet.get_value(row, col), (int, float))
        )
        return max_value if max_value is not None else float('-inf')

    @staticmethod
    def find_min(cell_range, sheet):
        """Find the minimum value within the given cell range."""
        if not FunctionLibrary.is_valid_range(cell_range, sheet):
            raise ValueError("Invalid cell range.")
        start, end = cell_range.split(":")
        start_row, start_col = Cell().cell_loc(start)
        end_row, end_col = Cell().cell_loc(end)

        min_value = min(
            sheet.get_value(row, col)
            for row in range(start_row, end_row + 1)
            for col in range(start_col, end_col + 1)
            if isinstance(sheet.get_value(row, col), (int, float))
        )

        return min_value if min_value is not None else float('inf')

