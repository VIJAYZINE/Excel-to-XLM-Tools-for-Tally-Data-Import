import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re

class ExcelToXMLConverter:
    def __init__(self, master):
        self.master = master
        master.title("Excel to XML Converter for Tally Data Import")

        # UI Elements
        self.excel_file_label = ttk.Label(master, text="Excel File:")
        self.excel_file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.excel_file_path = tk.StringVar()
        self.excel_file_entry = ttk.Entry(master, textvariable=self.excel_file_path, width=50)
        self.excel_file_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.browse_button = ttk.Button(master, text="Browse", command=self.browse_excel_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")

        self.invoice_format_label = ttk.Label(master, text="Invoice Format:")
        self.invoice_format_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.invoice_format_options = ["Sales Bill (Format 1)", "Purchase Entry (Format 1)", "Custom Format"]  # Add more formats as needed
        self.invoice_format = tk.StringVar(value=self.invoice_format_options[0])
        self.invoice_format_dropdown = ttk.Combobox(master, textvariable=self.invoice_format, values=self.invoice_format_options, state="readonly")
        self.invoice_format_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.custom_format_frame = ttk.Frame(master)  # Frame to hold custom format options
        self.custom_format_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        self.custom_format_frame.columnconfigure(1, weight=1)  # Make column 1 expandable

        self.custom_format_elements = {}  # Store mapping of Excel column to XML tag for custom formats
        self.show_hide_custom_format_options() #Initial Visibility

        self.invoice_format_dropdown.bind("<<ComboboxSelected>>", self.show_hide_custom_format_options)


        self.xml_tag_label = ttk.Label(master, text="XML Root Tag:")
        self.xml_tag_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")

        self.xml_tag = tk.StringVar(value="Vouchers")  # Default XML root tag
        self.xml_tag_entry = ttk.Entry(master, textvariable=self.xml_tag, width=30)
        self.xml_tag_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        self.output_file_label = ttk.Label(master, text="Output XML File:")
        self.output_file_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        self.output_file_path = tk.StringVar()
        self.output_file_entry = ttk.Entry(master, textvariable=self.output_file_path, width=50)
        self.output_file_entry.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

        self.browse_output_button = ttk.Button(master, text="Browse", command=self.browse_output_file)
        self.browse_output_button.grid(row=4, column=2, padx=5, pady=5, sticky="w")

        self.preview_button = ttk.Button(master, text="Preview Data", command=self.preview_data)
        self.preview_button.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

        self.convert_button = ttk.Button(master, text="Convert to XML", command=self.convert_to_xml)
        self.convert_button.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

        self.error_log_label = ttk.Label(master, text="Error Log:")
        self.error_log_label.grid(row=7, column=0, padx=5, pady=5, sticky="w")

        self.error_log_text = tk.Text(master, height=10, width=70, state="disabled")
        self.error_log_text.grid(row=8, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

        # Configure column weights for resizing
        master.columnconfigure(1, weight=1) # Make the middle column expandable

    def browse_excel_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        self.excel_file_path.set(filename)

    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")])
        self.output_file_path.set(filename)

    def show_hide_custom_format_options(self):
        selected_format = self.invoice_format.get()

        #Clear any existing widgets in the custom format frame
        for widget in self.custom_format_frame.winfo_children():
            widget.destroy()

        self.custom_format_elements = {} #Reset dictionary

        if selected_format == "Custom Format":
            #  Dynamically create fields for column-to-tag mapping
            #  These are just example column names. Adjust to match expected Excel structure
            example_columns = ["Invoice Number", "Date", "Party Name", "GSTIN", "Item Name", "Quantity", "Rate", "Amount"]  # Example Columns
            for i, col in enumerate(example_columns):
                label = ttk.Label(self.custom_format_frame, text=f"Excel Column '{col}':")
                label.grid(row=i, column=0, padx=5, pady=2, sticky="w")

                xml_tag_var = tk.StringVar()
                entry = ttk.Entry(self.custom_format_frame, textvariable=xml_tag_var, width=20)
                entry.grid(row=i, column=1, padx=5, pady=2, sticky="ew")  # Expand entry width
                self.custom_format_elements[col] = xml_tag_var  # Store StringVar

                required_var = tk.BooleanVar(value=True)  # By default, required
                required_check = ttk.Checkbutton(self.custom_format_frame, text="Required", variable=required_var)
                required_check.grid(row=i, column=2, padx=5, pady=2, sticky="w")
                self.custom_format_elements[col + "_required"] = required_var  # Store BooleanVar


        else:
            # Hide the custom format options if not selected.
            pass

    def preview_data(self):
        excel_file = self.excel_file_path.get()
        try:
            df = pd.read_excel(excel_file)
            # Display the DataFrame (e.g., in a new window or in the text area)
            preview_window = tk.Toplevel(self.master)
            preview_window.title("Data Preview")
            text_area = tk.Text(preview_window, wrap="word", height=20, width=80)
            text_area.pack(padx=10, pady=10)
            text_area.insert(tk.END, df.to_string())
            text_area.config(state=tk.DISABLED) #Make read only

        except FileNotFoundError:
            messagebox.showerror("Error", "Excel file not found.")
        except Exception as e:
            messagebox.showerror("Error", f"Error reading Excel file: {e}")

    def convert_to_xml(self):
        excel_file = self.excel_file_path.get()
        output_file = self.output_file_path.get()
        invoice_format = self.invoice_format.get()
        xml_root_tag = self.xml_tag.get()

        if not excel_file or not output_file:
            messagebox.showerror("Error", "Please select both Excel file and output XML file.")
            return

        try:
            df = pd.read_excel(excel_file)
            error_log = []  # Collect errors during validation and conversion

            root = ET.Element(xml_root_tag)  # Root element

            for index, row in df.iterrows():
                voucher = ET.SubElement(root, "Voucher")  # Individual Voucher element

                #  Data extraction and XML element creation based on selected format
                if invoice_format == "Sales Bill (Format 1)":
                    #  Example mapping - Adapt to your Excel layout and Tally XML format
                    self.create_element(voucher, "VoucherNumber", row.get("Invoice Number"), error_log)
                    self.create_element(voucher, "Date", row.get("Date"), error_log, is_date=True)
                    self.create_element(voucher, "PartyLedgerName", row.get("Party Name"), error_log)
                    # ... more fields

                elif invoice_format == "Purchase Entry (Format 1)":
                    # Different logic for Purchase Entry
                    self.create_element(voucher, "VoucherNumber", row.get("Purchase Invoice Number"), error_log)
                    self.create_element(voucher, "Date", row.get("Purchase Date"), error_log, is_date=True)
                    self.create_element(voucher, "SupplierName", row.get("Supplier Name"), error_log)

                elif invoice_format == "Custom Format":
                    #  Custom mapping based on user input

                    for excel_column, xml_tag_var in self.custom_format_elements.items():
                        if excel_column.endswith("_required"):
                            continue # Skip required flags
                        xml_tag = xml_tag_var.get()
                        required = self.custom_format_elements[excel_column + "_required"].get()

                        excel_column_name = excel_column # The actual excel column name
                        if xml_tag:  # Only process if a tag name is provided
                            self.create_element(voucher, xml_tag, row.get(excel_column_name), error_log, required=required)
                else:
                    error_log.append(f"Row {index + 2}: Invalid Invoice Format selected.")


            #  Write XML to file (using pretty printing for readability)
            xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="   ")
            with open(output_file, "w", encoding="utf-8") as f:  # UTF-8 encoding
                f.write(xml_str)


            # Display error log
            self.update_error_log(error_log)

            if error_log:
                messagebox.showwarning("Warning", "XML conversion completed with errors. See error log.")
            else:
                messagebox.showinfo("Success", "XML conversion completed successfully!")

        except FileNotFoundError:
            messagebox.showerror("Error", "Excel file not found.")
            self.update_error_log(["Excel file not found."])  # Log error
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.update_error_log([f"An error occurred: {e}"]) # Log error


    def create_element(self, parent, tag, value, error_log, is_date=False, required=True):
        """Creates an XML element and handles data validation/conversion."""
        if value is None or pd.isna(value): # Handle missing/NaN values correctly
            if required:
               error_log.append(f"Missing required value for '{tag}'.")
            return  # Skip if value is missing and not required

        try:
            if is_date:
                # Convert to Tally's date format (e.g., DD-MM-YYYY)
                if isinstance(value, pd.Timestamp):  #Handles date format from excel
                   value = value.strftime("%d-%m-%Y")
                else:  # If the excel column is already a string, try to handle the format
                   try:
                       value = pd.to_datetime(value).strftime("%d-%m-%Y") # convert from various date formats
                   except:
                       error_log.append(f"Invalid date format for '{tag}': {value}")
                       return  #Skip invalid date

            # Convert numeric types explicitly to string to avoid XML errors
            elif isinstance(value, (int, float)):
                value = str(value)
            else:
                value = str(value)  # Convert to string if not already

            element = ET.SubElement(parent, tag)
            element.text = value
        except Exception as e:
            error_log.append(f"Error creating element '{tag}': {e}")


    def update_error_log(self, errors):
        """Updates the error log text area in the GUI."""
        self.error_log_text.config(state="normal")  # Enable editing
        self.error_log_text.delete("1.0", tk.END)  # Clear previous content
        for error in errors:
            self.error_log_text.insert(tk.END, error + "\n")
        self.error_log_text.config(state="disabled")  # Disable editing


# Main application setup
root = tk.Tk()
converter = ExcelToXMLConverter(root)
root.mainloop()