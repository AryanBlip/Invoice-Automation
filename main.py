print("WELCOME TO INVOICE GENERATION !\n")
# CREATE AN INVOICE IN 40 SECONDS
from tkinter import Tk, Toplevel, TclError
from tkinter import messagebox, Label
from tkinter import ttk, StringVar
from tkinter import filedialog, Button

#TODO: ADD INSTRUCTIONS
print("PRE REQUISITES >")
print("1> Copy the tabular data into an Excel file ALONG WITH ITS HEADER ROW")
print("2> Save the Excel file \n")

print("STEPS TO CREATE INVOICE:-")
print("Step 1: Select the bank in the 'Select Bank' dropdown")
print("Step 2: Click 'Browse..' to select the saved Excel file - Click '*All Files' on Bottom Right to view all files")
print("Step 3: Click 'Load Data' to load data from the Excel file")
print("Step 4: Ensure correct information/calculation - Double click to edit a value and click 'Enter' on Keyboard")
print("Step 5: Enter the Month and Year in the 'Month Year'")
print("Step 6: Click 'Create Invoice' to proceed")
print("Step 7: Select the directory to save the newly created invoice")
print("Step 8: Wait for confirmation Pop Up")
print("Step 9: Find your newly saved invoice")

class MainApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("Invoice Generation - Bank Selection")
        self.root.geometry("400x250") # Adjust window size to fit new elements

        #TODO: ADD BANKS HERE
        self.bill_to_banks = {
            'ADIB': 1,
            'DIB': 2,
            'CBD' : 3,
            'MASHREQ' : 4
        }

        # New attribute to store the selected Excel file path
        self.excel_file_path = None

        # Bank Selection Widgets
        label_bank = Label(self.root, text="Select a Bank:")
        label_bank.pack(pady=10)

        self.selected_bank = StringVar(self.root)
        self.selected_bank.set("Select Bank")

        bank_options = tuple(self.bill_to_banks.keys())
        dropdown = ttk.Combobox(self.root, textvariable=self.selected_bank, values=bank_options, state="readonly")
        dropdown.pack(pady=5)
        
        # New widgets for Excel file selection
        label_file = Label(self.root, text="Select an Excel File:")
        label_file.pack(pady=10)

        self.file_path_label = Label(self.root, text="No file selected", fg="gray")
        self.file_path_label.pack(pady=5)

        select_file_btn = Button(self.root, text="Browse...", command=self.ask_excel_file)
        select_file_btn.pack(pady=5)

        # Button to start the process
        start_btn = Button(self.root, text="Load Data", command=self.start_automation)
        start_btn.pack(pady=20)
        
    def ask_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Select an Excel or ODF file",
            filetypes=(
                ("Spreadsheet files", "*.xlsx *.ods"), 
                ("Excel files", "*.xlsx"), 
                ("ODF files", "*.ods"), 
                ("All files", "*.*")
            )
        )
        if file_path:
            self.excel_file_path = file_path
            self.file_path_label.config(text=f"File: {self.excel_file_path.split('/')[-1]}", fg="green")
        else:
            self.excel_file_path = None
            self.file_path_label.config(text="No file selected", fg="gray")

    def start_automation(self):
        selected_bank_name = self.selected_bank.get()

        # TODO: ADD BANKS HERE

        if selected_bank_name == "ADIB":
            import adib_module as adib

            self.bill_to_banks = {
                'ADIB': adib.InvoiceAutomation,
            }
            
        elif selected_bank_name == "DIB":
            import dib_module as dib    

            self.bill_to_banks = {
                'DIB': dib.InvoiceAutomation
            }

        elif selected_bank_name == "CBD":
            import cbd_module as cbd

            self.bill_to_banks = {
                'CBD': cbd.InvoiceAutomation
            }
        
        elif selected_bank_name == "MASHREQ":
            import mashreq_module as msq

            self.bill_to_banks = {
                'MASHREQ' : msq.InvoiceAutomation
            }

        print(f"Processing file : {self.excel_file_path}\n")

        # Check if both a bank and a file have been selected
        if selected_bank_name in self.bill_to_banks and self.excel_file_path:
            bank_app_class = self.bill_to_banks[selected_bank_name]
            
            # Pass the file path to the bank's automation class
            bank_app = bank_app_class(Toplevel(self.root), self.excel_file_path, self.root)

            try:
                self.root.withdraw()
                bank_app.run()
            except KeyboardInterrupt:
                print("Exiting Bank.")
            except TclError:
                print("\nRe-Starting Main Application...")
                app = MainApp()
                app.run()
        else:
            messagebox.showerror("Error", "Please select both a valid bank and an Excel file.")

    def run(self):
        try : 
            self.root.mainloop()
        except KeyboardInterrupt:
            print("Application Stopped.")

if __name__ == "__main__":
    try:
        app = MainApp()
        app.run()
        print("\nWindow can be closed.\nAuto Closing in 2 Minutes...")

    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        from time import sleep
        sleep(120)