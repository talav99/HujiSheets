from sheet import Sheet
import pytest
from utils.calculates import FormulaParser, FunctionLibrary

@pytest.fixture
def sheet():
    return Sheet(3, 3)  # Create a sample sheet for testing with 3 rows and 3 columns

def test_set_value(sheet):
    sheet.set_value(5, 0, 0)
    assert sheet.get_value(0, 0) == 5

def test_set_formula(sheet):
    sheet.set_formula("=A1+B1", 1, 1)
    assert sheet.get_formula(1, 1) == "=A1+B1"

def test_get_value_out_of_range(sheet):
    with pytest.raises(IndexError):
        sheet.get_value(5, 5)

def test_set_value_out_of_range(sheet):
    with pytest.raises(IndexError):
        sheet.set_value(10, 5, 5)

def test_set_formula_out_of_range(sheet):
    with pytest.raises(IndexError):
        sheet.set_formula("=A1+B1", 5, 5)

def test_get_formula_cells(sheet):
    sheet.set_formula("=A1+B1", 1, 1)
    formula_cells = sheet.get_formula_cells()
    assert formula_cells == {"B2": "=A1+B1"}

def test_get_formula_cells_multiple_formulas(sheet):
    sheet.set_formula("=A1+B1", 1, 1)
    sheet.set_formula("=A2-B2", 2, 2)
    formula_cells = sheet.get_formula_cells()
    assert formula_cells == {"B2": "=A1+B1", "C3": "=A2-B2"}

@pytest.fixture
def test_sheet1():
    # Create a test sheet instance with predefined values
    sheet = Sheet(5, 5)
    # Set some test values
    sheet.set_value(10, 0, 0)
    sheet.set_value(20, 0, 1)
    sheet.set_value(30, 1, 0)
    sheet.set_value(40, 1, 1)
    return sheet

def test_evaluate_formula_sum(test_sheet1):
    result = FormulaParser.evaluate_formula("SUM(A1:B2)", test_sheet1)
    assert result == 100

def test_evaluate_formula_average(test_sheet1):
    result = FormulaParser.evaluate_formula("AVERAGE(A1:B2)", test_sheet1)
    assert result == 25

def test_evaluate_formula_max(test_sheet1):
    result = FormulaParser.evaluate_formula("MAX(A1:B2)", test_sheet1)
    assert result == 40

def test_evaluate_formula_min(test_sheet1):
    result = FormulaParser.evaluate_formula("MIN(A1:B2)", test_sheet1)
    assert result == 10

def test_calc_sum(test_sheet1):
    result = FunctionLibrary.calc_sum("A1:B2", test_sheet1)
    assert result == 100

def test_find_max(test_sheet1):
    result = FunctionLibrary.find_max("A1:B2", test_sheet1)
    assert result == 40

def test_find_min(test_sheet1):
    result = FunctionLibrary.find_min("A1:B2", test_sheet1)
    assert result == 10

@pytest.fixture
def test_sheet2():
    # Create a test sheet instance with predefined values
    sheet = Sheet(5, 5)
    # Fill some cells with values
    for i in range(5):
        for j in range(5):
            sheet.set_value(i * 10 + j, i, j)
    return sheet

def test_evaluate_formula_average1(test_sheet2):
    # Test the average calculation
    result = FormulaParser.evaluate_formula('=average(A1:E5)', test_sheet2)
    expected_average = (0 + 1 + 2 + 3 + 4 + 10 + 11 + 12 + 13 + 14 + 20 + 21 + 22 + 23 + 24 +
                        30 + 31 + 32 + 33 + 34 + 40 + 41 + 42 + 43 + 44) / 25
    assert result == expected_average


def test_convert_to_math_formula():
    test_sheet = Sheet(3, 3)
    test_sheet.set_value(5, 0, 0)
    test_sheet.set_value(10, 0, 1)

    # Test addition
    assert FormulaParser.convert_to_math_formula("A1 + B1", test_sheet) == 15
    # Test subtraction
    assert FormulaParser.convert_to_math_formula("A1 - B1", test_sheet) == -5
    # Test multiplication
    assert FormulaParser.convert_to_math_formula("A1 * B1", test_sheet) == 50
    # Test division
    assert FormulaParser.convert_to_math_formula("A1 / B1", test_sheet) == 0.5
    # Test division by zero
    with pytest.raises(ValueError, match="Division by zero error!"):
        FormulaParser.convert_to_math_formula("A1 / B2", test_sheet)

def test_find_cell_len():
    # Test range with 2x2 cells
    assert FunctionLibrary.find_cell_len("A1:B2") == 4
    # Test range with 1x3 cells
    assert FunctionLibrary.find_cell_len("A1:C1") == 3


def test_average():
    test_sheet = Sheet(3, 3)
    test_sheet.set_value(5, 0, 0)
    test_sheet.set_value(10, 0, 1)
    test_sheet.set_value(15, 1, 0)
    test_sheet.set_value(20, 1, 1)
    # Test average calculation
    assert FunctionLibrary.average("A1:B2", test_sheet) == 12.5
    # Test average of empty range
    assert FunctionLibrary.average("A3:B3", test_sheet) == 0


def test_find_max():
    test_sheet = Sheet(3, 3)
    test_sheet.set_value(5, 0, 0)
    test_sheet.set_value(10, 0, 1)
    test_sheet.set_value(15, 1, 0)
    test_sheet.set_value(20, 1, 1)
    # Test maximum value calculation
    assert FunctionLibrary.find_max("A1:B2", test_sheet) == 20

def test_find_min():
    test_sheet = Sheet(3, 3)
    test_sheet.set_value(5, 0, 0)
    test_sheet.set_value(10, 0, 1)
    test_sheet.set_value(15, 1, 0)
    test_sheet.set_value(20, 1, 1)
    # Test minimum value calculation
    assert FunctionLibrary.find_min("A1:B2", test_sheet) == 5
