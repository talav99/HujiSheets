from utils.sheet import Sheet, Cell
import tkinter as tk
import tkinter.simpledialog as simpledialog
from tkinter import messagebox, filedialog
import tkinter.colorchooser as colorchooser
from utils.calculates import FormulaParser
import string
import sys
from utils.import_export import SheetLoader

window = tk.Tk()
window.geometry("800x500")
window.title("sheet")

class UserInterface:
    def __init__(self):
        self.file_path = None
        self.entry_font_size = 10
        self.cell_colors = {}
        self.entry_boxes = {}
        self.font_colors = {}
        # Ask the user whether to import a file or create a new sheet
        self.choose_option()

    def configure_menu_bar(self):
        menu_bar = tk.Menu(window)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Import JSON/YAML", command=self.load_file)
        file_menu.add_separator()
        file_menu.add_command(label="Export to JSON", command=lambda: self.export_to_file('.json'))
        file_menu.add_command(label="Export to YAML", command=lambda: self.export_to_file('.yaml'))
        file_menu.add_command(label="Export to PDF", command=self.export_to_pdf)
        file_menu.add_command(label="Export to CSV", command=self.export_to_csv)
        file_menu.add_command(label="Export to Excel", command=self.export_to_excel)

        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_application)

        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_separator()
        edit_menu.add_command(label="Change Cell Color", command=self.prompt_color_change)
        edit_menu.add_command(label="Change Font Color", command=self.prompt_font_color_change)
        edit_menu.add_separator()
        edit_menu.add_command(label="Increase Font Size", command=self.increase_font_size)
        edit_menu.add_command(label="Decrease Font Size", command=self.decrease_font_size)

        menu_bar.add_cascade(label="File", menu=file_menu)
        menu_bar.add_cascade(label="Edit", menu=edit_menu)

        window.config(menu=menu_bar)

    def choose_option(self):
        option = messagebox.askquestion("Choose Option", "Would you like to import"
                                                         "a file or create a new sheet?"
                                                         "\n press Yes to import, No to create a new sheet")
        if option == 'yes':
            self.load_file()
        else:
            self.prompt_sheet_dimensions()


    def update_formula_cells(self):
        formula_cells = self.sheet.get_formula_cells()
        for (row, col), entry in self.entry_boxes.items():
            cell_address = self.sheet.get_cell_address(row, col)
            if cell_address in formula_cells:
                formula = entry.get()
                if formula.startswith('='):
                    # Remove leading '=' and whitespace
                    formula = formula[1:].strip()
                    if not formula.startswith('clr'):
                        result = FormulaParser.evaluate_formula(formula, self.sheet)
                        if result is not None:
                            self.sheet.set_value(result, row, col)
                            entry.delete(0, tk.END)
                            entry.insert(0, str(result))  # Update entry text with result

    def bind_enter_key(self, entry):
        # Bind the Enter key to the entry widget
        if entry.winfo_exists():
            entry.bind('<Return>', lambda event: self.move_focus_down(entry))

    def bind_focus_out_event(self, entry):
        # Bind the <FocusOut> event to update the cell when the user exits the entry field
        if entry.winfo_exists():
            entry.bind('<FocusOut>', lambda event: self.update_formula_cells())
        entry.bind('<FocusIn>', lambda event: self.recalculate_all_formulas())

    def move_focus_down(self, entry):
        current_row, current_col = self.get_entry_position(entry)
        next_row = current_row + 1
        next_col = current_col
        if next_row < self.sheet.rows:
            next_entry = self.entry_boxes[(next_row, next_col)]
            next_entry.focus_set()

    def get_entry_position(self, entry):
        """Find the position (row, col) of the given entry widget"""
        for (row, col), widget in self.entry_boxes.items():
            if widget == entry:
                return row, col
        return -1, -1  # Default if not found

    def set_cell_color(self, row, col, color):
        """Set the color of the cell at the specified row and column."""
        cell_address = self.sheet.get_cell_address(row, col)
        self.cell_colors[cell_address] = color

    def get_cell_color(self, row, col):
        """Get the color of the cell at the specified row and column."""
        cell_address = self.sheet.get_cell_address(row, col)
        return self.cell_colors.get(cell_address, '#FFFFFF')  # Default color is black

    def get_cell_colors(self):
        """Get the cell colors as a 2D array."""
        cell_colors = []
        for row_index in range(self.sheet.rows):
            row_colors = []
            for col_index in range(self.sheet.cols):
                cell_color = self.get_cell_color(row_index, col_index)
                row_colors.append(cell_color)
            cell_colors.append(row_colors)
        return cell_colors

    def set_font_color(self, row, col, color):
        """Set the font color of the cell at the specified row and column."""
        cell_address = self.sheet.get_cell_address(row, col)
        self.font_colors[cell_address] = color

    def get_font_color(self, row, col):
        """Get the font color of the cell at the specified row and column."""
        cell_address = self.sheet.get_cell_address(row, col)
        return self.font_colors.get(cell_address, '#000000')  # Default font color is black

    def prompt_sheet_dimensions(self):
        """Prompt from user sheet dimensions for a new sheet"""
        while True:
            try:
                dimensions = simpledialog.askstring("Enter Sheet Size",
                                                    "Enter the number of rows "
                                                    "(1-50) and columns (1-25) separated by comma:")
                if dimensions:
                    rows, cols = map(int, dimensions.split(","))
                    if 1 <= rows <= 50 and 1 <= cols <= 25:
                        self.sheet = Sheet(rows, cols)
                        self.display_sheet()
                        break
                    else:
                        messagebox.showerror("Error", "Invalid dimensions!"
                                                      "Enter values within the specified range.")
            except ValueError:
                messagebox.showerror("Error", "Please enter comma-separated integers only.")

    def display_sheet(self):
        """ Code to display the spreadsheet in the GUI"""
        label_row = tk.Frame(window)
        for col_index, label in enumerate(['']):
            entry_width = 10  # Default width
            tk.Label(label_row, text=label, font=("Arial", 10), width=entry_width).grid(
                row=0, column=col_index, padx=2, pady=2)
        label_row.grid(row=0, column=0, columnspan=self.sheet.cols + 1, sticky="nsew")

        # Add non-editable entry widgets for column indices on top of the first row
        for col_index in range(self.sheet.cols):
            entry_width = self.get_entry_width()  # Use the width of existing entry widgets
            entry_text = string.ascii_uppercase[col_index]  # Get the corresponding column label (A, B, C, ...)
            entry = tk.Entry(window, font=("Arial", 10), width=entry_width)
            entry.insert(0, entry_text)  # Insert column label
            entry.config(state='readonly')  # Set entry to read-only
            entry.grid(row=0, column=col_index + 1, padx=2, pady=2, sticky="nsew")

        # Add entry widgets for cells
        for row_index in range(self.sheet.rows):
            # Add label for row index
            row_label = tk.Label(window, text=str(row_index + 1), font=("Arial", 10), width=3)  # Adjust width here
            row_label.grid(row=row_index + 1, column=0, padx=2, pady=2, sticky="nsew")

            for col_index in range(self.sheet.cols):
                value = self.sheet.get_value(row_index, col_index)
                formula = self.sheet.get_formula(row_index, col_index)

                entry = tk.Entry(window, font=("Arial", 10), borderwidth=1, relief="raised", width=10)
                entry.insert(0, str(value) if value is not None else str(formula))
                entry.grid(row=row_index + 1, column=col_index + 1, padx=2, pady=2, sticky="nsew")
                self.bind_focus_out_event(entry)  # Bind FocusOut event

                # Bind the <FocusOut> event to update the cell when the user exits the entry field
                entry.bind('<FocusOut>',
                           lambda event, row=row_index, col=col_index, entry=entry: self.update_cell(row, col, entry))

                # Bind Enter key for all entry widgets
                self.bind_enter_key(entry)
                self.entry_boxes[(row_index, col_index)] = entry  # Store the entry widget after creation

        # Configure column and row weights for dynamic resizing
        window.grid_columnconfigure(0, weight=1)  # Row index column
        for i in range(self.sheet.cols):
            window.grid_columnconfigure(i + 1, weight=1)
        for i in range(self.sheet.rows):
            window.grid_rowconfigure(i + 1, weight=1)

    def get_entry_width(self):
        """Determine the width of the existing entry widgets."""
        max_width = 0
        for entry in self.entry_boxes.values():
            width = entry.winfo_width()
            if width > max_width:
                max_width = width
        return max_width


    def update_cell(self, row, col, entry):
        try:
            new_value = entry.get()  # Get the new value from the entry widget
            if new_value.startswith('='):
                formula = new_value[1:].strip()  # Remove leading '='
                if formula.upper().startswith(('SUM', 'AVERAGE', 'MAX', 'MIN')):
                    # Evaluate built-in function
                    result = FormulaParser.evaluate_formula(formula, self.sheet)
                elif formula.startswith('clr'):
                    # Apply color formula
                    self.apply_color_formula(formula[4:-1])  # Extract cell range from formula
                    result = None
                else:
                    # Evaluate the mathematical expression
                    result = FormulaParser.convert_to_math_formula(formula, self.sheet)

                if result is not None:
                    self.sheet.set_value(result, row, col)
                    self.sheet.set_formula(new_value, row, col)
                    entry.delete(0, tk.END)
                    entry.insert(0, str(result))

                    # Reevaluate all formula cells
                    self.recalculate_all_formulas()
            else:
                self.sheet.set_value(new_value, row, col)
                # Update entry text to match the new value
                entry.delete(0, tk.END)
                entry.insert(0, str(new_value))

            # Check for division by zero errors in formula cells
            self.check_for_division_by_zero()

        except ValueError as e:
            # Handle invalid input (e.g., non-numeric values)
            messagebox.showerror("Error", str(e))
            print("Error:", e)

    def check_for_division_by_zero(self):
        try:
            self.recalculate_all_formulas()
        except ZeroDivisionError:
            # Handle division by zero error
            messagebox.showerror("Error", "Division by zero error detected in formula.")
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            print("Error:", e)

    def recalculate_all_formulas(self):
        for (row, col), entry in self.entry_boxes.items():
            formula = self.sheet.get_formula(row, col)
            if formula is not None and formula.startswith('='):
                # Remove leading '=' and whitespace
                formula = formula[1:].strip()
                if formula.upper().startswith(('SUM', 'AVERAGE', 'MAX', 'MIN')):
                    # Evaluate built-in function
                    result = FormulaParser.evaluate_formula(formula, self.sheet)
                elif formula.startswith('clr'):
                    # Apply color formula
                    self.apply_color_formula(formula[4:-1])  # Extract cell range from formula
                    result = None
                else:
                    # Evaluate the mathematical expression
                    result = FormulaParser.convert_to_math_formula(formula, self.sheet)

                if result is not None:
                    self.sheet.set_value(result, row, col)
                    entry.delete(0, tk.END)
                    entry.insert(0, str(result))  # Update entry text with result

    def load_file(self):
        """load imported file"""
        file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json"), ("YAML Files", "*.yaml")])
        if file_path:
            self.file_path = file_path
            if self.file_path.endswith('.json'):
                loaded_sheet = SheetLoader.load_from_json(self.file_path)
            elif self.file_path.endswith('.yaml'):
                loaded_sheet = SheetLoader.load_from_yaml(self.file_path)
            else:
                messagebox.showerror("Error", "Unsupported file format.")
                return

            if loaded_sheet:
                self.sheet = loaded_sheet
                self.display_sheet()

    def prompt_font_color_change(self):
        """Prompt the user to choose a font color for the selected cell"""
        cell_address = self.get_selected_cell()
        if cell_address:
            color = colorchooser.askcolor(title="Choose Font Color")
            if color[1]:  # If a color was chosen
                self.set_font_color(cell_address[0], cell_address[1], color[1])
                self.refresh_font_colors()

    def refresh_font_colors(self):
        """Update font colors of all cells based on sheet data"""
        for (row, col), entry in self.entry_boxes.items():
            font_color = self.get_font_color(row, col)
            entry.config(fg=font_color)

    def prompt_color_change(self):
        """Allow user to choose a color for the selected cell"""
        cell_address = self.get_selected_cell()
        if cell_address:
            color = colorchooser.askcolor(title="Choose Color")
            if color[1]:  # If a color was chosen
                self.set_cell_color(cell_address[0], cell_address[1], color[1])
                self.refresh_cell_colors()

    def refresh_cell_colors(self):
        for (row, col), entry in self.entry_boxes.items():
            cell_color = self.get_cell_color(row, col)
            entry.config(bg=cell_color)

    def get_selected_cell(self):
        focused_widget = window.focus_get()
        for (row, col), entry in self.entry_boxes.items():
            if entry == focused_widget:
                return row, col
        return None

    def apply_color_formula(self, cell_range):
        start, end = cell_range.split(":")
        start_row, start_col = Cell().cell_loc(start)
        end_row, end_col = Cell().cell_loc(end)
        # Prompt the user to choose a color
        color = colorchooser.askcolor(title="Choose Color")
        if color[1]:  # If a color was chosen
            chosen_color = color[1]
            # Apply color to each cell in the range
            for row in range(start_row, end_row + 1):
                for col in range(start_col, end_col + 1):
                    self.set_cell_color(row, col, chosen_color)
        self.refresh_cell_colors()

    def increase_font_size(self):
        self.entry_font_size += 1
        self.refresh_entry_font_size()

    def decrease_font_size(self):
        if self.entry_font_size > 1:
            self.entry_font_size -= 1
            self.refresh_entry_font_size()

    def refresh_entry_font_size(self):
        for entry in self.entry_boxes.values():
            entry.config(font=("Arial", self.entry_font_size))

    def exit_application(self):
        window.destroy()

    def export_to_file(self, file_extension):
        if not self.sheet:
            messagebox.showerror("Error", "No sheet data to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=file_extension,
                                                 filetypes=[("JSON Files", "*.json"), ("YAML Files", "*.yaml")])
        if file_path:
            if file_extension == '.json':
                SheetLoader.export_to_json(self.sheet, file_path)
            elif file_extension == '.yaml':
                SheetLoader.export_to_yaml(self.sheet, file_path)

    def export_to_pdf(self):
        """Export the sheet to a PDF file."""
        if not self.sheet:
            messagebox.showerror("Error", "No sheet data to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            try:
                cell_colors = self.get_cell_colors()  # Get cell colors
                SheetLoader.export_to_pdf(self.sheet, file_path, cell_colors)  # Pass cell colors
                messagebox.showinfo("Success", f"Sheet exported to PDF: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

    def export_to_csv(self):
        """Export the sheet to a CSV file."""
        if not self.sheet:
            messagebox.showerror("Error", "No sheet data to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                SheetLoader.export_to_csv(self.sheet, file_path)
                messagebox.showinfo("Success", f"Sheet exported to CSV: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

    def export_to_excel(self):
        """Export the sheet to an Excel (XLSX) file."""
        if not self.sheet:
            messagebox.showerror("Error", "No sheet data to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                SheetLoader.export_to_excel(self.sheet, file_path)
                messagebox.showinfo("Success", f"Sheet exported to Excel: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")

def main():
    user_interface = UserInterface()
    user_interface.configure_menu_bar()
    window.mainloop()

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--help":
        print("call 'main' to activate the program. Choose Sheet dimensions or import a file. Then,"
              "you will be able to edit your sheet")
    else:
        main()
