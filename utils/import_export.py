import json
from sheet import Sheet
import yaml
from openpyxl import Workbook
import csv
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

class SheetLoader:
    @staticmethod
    def load_from_json(file_path):
        """Load sheet from imported json file"""
        try:
            with open(file_path, 'r') as file:
                data = json.load(file)
                rows = data.get('rows', 0)
                cols = data.get('cols', 0)
                # Create a new Sheet object with the appropriate dimensions
                sheet = Sheet(rows, cols)

                # Iterate over the data and set values in the sheet
                for cell_data in data.get('cells', []):
                    row_index = cell_data.get('row', 0)
                    col_index = cell_data.get('col', 0)
                    value = cell_data.get('value')
                    sheet.set_value(value, row_index, col_index)
                return sheet
        except FileNotFoundError:
            print(f"Error: File '{file_path}' not found.")
        except json.JSONDecodeError:
            print(f"Error: Failed to decode JSON file '{file_path}'.")
        except Exception as e:
            print(f"Error: An unexpected error occurred: {str(e)}")

    @staticmethod
    def export_to_json(sheet, file_path):
        """Export sheet to a json file"""
        data = {"rows": sheet.rows, "cols": sheet.cols, "cells": []}
        for row_index in range(sheet.rows):
            for col_index in range(sheet.cols):
                cell_data = {"row": row_index, "col": col_index, "value": sheet.get_value(row_index, col_index)}
                data["cells"].append(cell_data)
        with open(file_path, 'w') as file:
            json.dump(data, file, indent=4)

    @staticmethod
    def load_from_yaml(file_path):
        """Load sheet from a yaml file"""
        try:
            with open(file_path, 'r') as file:
                data = yaml.safe_load(file)
                rows = data.get('rows', 0)
                cols = data.get('cols', 0)
                # Create a new Sheet object with the appropriate dimensions
                sheet = Sheet(rows, cols)
                # Iterate over the data and set values in the sheet
                for cell_data in data.get('cells', []):
                    row_index = cell_data.get('row', 0)
                    col_index = cell_data.get('col', 0)
                    value = cell_data.get('value')
                    sheet.set_value(value, row_index, col_index)
                return sheet
        except FileNotFoundError:
            print(f"Error: File '{file_path}' not found.")
        except yaml.YAMLError as e:
            print(f"Error: Failed to load YAML file '{file_path}': {str(e)}")
        except Exception as e:
            print(f"Error: An unexpected error occurred: {str(e)}")

    @staticmethod
    def export_to_yaml(sheet, file_path):
        """Export a sheet to a yaml file"""
        data = {"rows": sheet.rows, "cols": sheet.cols, "cells": []}
        for row_index in range(sheet.rows):
            for col_index in range(sheet.cols):
                cell_data = {"row": row_index, "col": col_index, "value": sheet.get_value(row_index, col_index)}
                data["cells"].append(cell_data)
        with open(file_path, 'w') as file:
            yaml.dump(data, file)

    @staticmethod
    def export_to_pdf(sheet, file_path, cell_colors):
        """Export sheet to a PDF file"""
        try:
            # Create a new PDF document
            doc = SimpleDocTemplate(file_path, pagesize=letter)
            elements = []

            # Convert sheet data and cell colors to a list of lists (2D arrays)
            data = []
            for row_index in range(sheet.rows):
                row_data = []
                for col_index in range(sheet.cols):
                    value = sheet.get_value(row_index, col_index)
                    row_data.append(value)
                data.append(row_data)

            # Create a table from the sheet data
            table = Table(data)

            # Define table style with cell colors
            style = TableStyle([('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black)])

            for row_index, row_colors in enumerate(cell_colors):
                for col_index, cell_color in enumerate(row_colors):
                    if cell_color:
                        style.add('BACKGROUND', (col_index, row_index), (col_index, row_index), cell_color)

            table.setStyle(style)
            elements.append(table)

            # Build PDF document
            doc.build(elements)

            print(f"Sheet exported to PDF: {file_path}")
        except Exception as e:
            print(f"Error: An unexpected error occurred: {str(e)}")

    @staticmethod
    def export_to_csv(sheet, file_path):
        """Export sheet to a CSV file"""
        try:
            with open(file_path, 'w', newline='') as csvfile:
                csvwriter = csv.writer(csvfile)
                for row_index in range(sheet.rows):
                    row_data = [sheet.get_value(row_index, col_index) for col_index in range(sheet.cols)]
                    csvwriter.writerow(row_data)
            print(f"Sheet exported to CSV: {file_path}")
        except Exception as e:
            print(f"Error: An unexpected error occurred: {str(e)}")

    def export_to_excel(sheet: Sheet, file_path):
        """Export sheet to an Excel (XLSX) file"""
        try:
            wb = Workbook()
            ws = wb.active

            for row_index in range(sheet.rows):
                for col_index in range(sheet.cols):
                    value = sheet.get_value(row_index, col_index)
                    ws.cell(row=row_index + 1, column=col_index + 1).value = value

            wb.save(file_path)
            print(f"Sheet exported to Excel: {file_path}")
        except Exception as e:
            print(f"Error: An unexpected error occurred: {str(e)}")