import tkinter
from tkinter import filedialog
import os
import json
import Google_Service
import csv


class GUI:
    def __init__(self):

        # gui elements
        self.window = tkinter.Tk(screenName="BCTool")
        self.window.title("BCTool")
        self.frame = tkinter.Frame()
        self.spreadsheet_box = tkinter.Listbox(self.frame)
        self.sheet_box = tkinter.Listbox(self.frame)
        # var items
        self.sv_bank_column = tkinter.StringVar(self.frame)
        self.sv_bank_date_column = tkinter.StringVar(self.frame)
        self.sv_bank_desc_column = tkinter.StringVar(self.frame)
        self.sv_sheet_column = tkinter.StringVar(self.frame)
        self.sheet_column = tkinter.Entry(self.frame, width=5, textvariable=self.sv_sheet_column)
        self.bank_column_frame = tkinter.Frame(self.frame)
        self.bank_column = tkinter.Entry(self.bank_column_frame, width=5, textvariable=self.sv_bank_column)
        self.bank_date_column = tkinter.Entry(self.bank_column_frame, width=5, textvariable=self.sv_bank_date_column)
        self.bank_desc_column = tkinter.Entry(self.bank_column_frame, width=5, textvariable=self.sv_bank_desc_column)
        self.result_box = tkinter.Listbox(self.frame, width=90, height=20)
        self.list_sheet_download = tkinter.Button(self.frame, text="Download Sheets")
        self.spreadsheet_download = tkinter.Button(self.frame, text="Download Sheet")
        self.enter_button = tkinter.Button(self.frame, text="Match", state="disabled")

        self.transactions_from_spreadsheet = []
        self.transactions_from_bank = []
        self.all_spreadsheets = []
        self.all_sheets = []
        self.file = None
        self.google = Google_Service.GoogleService()

        self.is_csv_set = False
        self.is_sheet_col_set = False
        self.is_sheet_set = False
        self.is_column_set = False

    def download_one_sheet(self):
        spreadsheet_index = self.spreadsheet_box.curselection()
        self.sheet_box.delete(0, "end") # clear the box
        if len(spreadsheet_index) == 1:
            self.google.download(spreadsheet_index[0])
            sheets = self.google.get_sheets()
            for i in range(len(sheets)):
                self.sheet_box.insert(i, sheets[i])

    def initialize_gui(self):
        # self.window.geometry("800x600")
        self.frame.pack()

        label = tkinter.Label(self.frame, text="Your Spreadsheets:")
        label.grid(column=0, row=0)

        scroll = tkinter.Scrollbar(self.frame)
        scroll["command"] = self.spreadsheet_box.yview
        self.spreadsheet_box.grid(column=0, row=1)
        scroll.grid(column=0, row=1, sticky=tkinter.N+tkinter.S+tkinter.E)
        self.spreadsheet_box["yscrollcommand"] = scroll.set

        self.list_sheet_download['command'] = self.handle_button_press
        self.list_sheet_download.grid(column=0, row=2)
        self.list_sheet_download["width"] = 15

        self.spreadsheet_download.grid(column=0, row=3)
        self.spreadsheet_download["width"] = 15
        self.spreadsheet_download["state"] = "disabled"
        self.spreadsheet_download["command"] = self.download_one_sheet

        sheets_label = tkinter.Label(self.frame, text="Sheets:")
        sheets_label.grid(column=1, row=0)

        self.sheet_box.grid(column=1, row=1)
        self.sheet_box.bind("<<ListboxSelect>>", self.update_compute_button)

        label_column = tkinter.Label(self.frame, text="Google Column:")
        label_column.grid(column=1, row=2)

        self.window.bind("<Key>", self.update_compute_button)
        self.sheet_column.grid(column=1, row=3)


        self.bank_column_frame.grid(column=2, row=1)

        label_bank_column = tkinter.Label(self.bank_column_frame, text="Bank Money Column")
        label_bank_column.grid(column=0, row=0)
        self.bank_column.grid(column=0, row=1)

        label_desc_column = tkinter.Label(self.bank_column_frame, text="Bank Description Column")
        label_desc_column.grid(row=2)
        self.bank_desc_column.grid(row=3)

        label_date_column = tkinter.Label(self.bank_column_frame, text="Bank Date Column")
        label_date_column.grid(row=4)
        self.bank_date_column.grid(row=5)


        label_bad = tkinter.Label(self.frame, text="Missing Transactions:")
        label_bad.grid(column=0, row=6)
        result_scroll = tkinter.Scrollbar(self.frame, orient=tkinter.VERTICAL, command=self.result_box.yview)
        self.result_box.configure(yscrollcommand=result_scroll.set)
        result_scroll.grid(row=7, column=0, columnspan=4, sticky=tkinter.N+tkinter.S+tkinter.E)
        self.result_box.grid(row=7, column=0,  columnspan=4)
        self.enter_button["command"] = self.compute
        self.enter_button.grid(column=1, row=5)

        browse_csv = tkinter.Button(self.frame, text="Browse for CSV...",
                                    command= self.browse)
        browse_csv.grid(column=2, row=4)

    def handle_button_press(self):
        self.list_sheet_download["state"] = "disabled"
        self.google.prepare_google_client()
        self.all_spreadsheets = self.google.get_drive_spreadsheets()
        self.spreadsheet_download["state"] = "active"
        for file in self.all_spreadsheets["files"]:
            self.spreadsheet_box.insert("end", file["name"])

    def browse(self):
        self.transactions_from_bank = []  # reset the value so that we don't get leaks.
        self.file = tkinter.filedialog.askopenfile(title="Select Transaction Document", defaultextension="csv")
        if self.file:
            self.is_csv_set = True
        else:
            self.is_csv_set = False

        self.update_compute_button()

    def save_configuration(self):
        save = {
            "google_sheet_column": self.sheet_column.get(),
            "bank_amount_column": self.bank_column.get(),
            "bank_date_column": self.bank_date_column.get(),
            "bank_desc_column": self.bank_desc_column.get(),
        }

        with open("config.json", "w") as json_file:
            json_file.write(json.dumps(save))

    def load_configuration(self):
        if os.path.exists("config.json"):
            with open("config.json", "r") as jsonfile:
                config = json.loads(jsonfile.read())

            self.sv_sheet_column.set(config["google_sheet_column"])
            self.sv_bank_column.set(config["bank_amount_column"])
            self.sv_bank_date_column.set(config["bank_date_column"])
            self.sv_bank_desc_column.set(config["bank_desc_column"])



    def compute(self):
        # first, get which sheet is selected.
        self.save_configuration()
        sheet_index = self.sheet_box.curselection()[0]

        sheet = self.google.spreadsheet_object["sheets"][sheet_index]

        # we have the sheet, let's determine what line the amount will be on...
        sheet_column = Google_Service.letter_to_number(self.sheet_column.get())
        bank_column = Google_Service.letter_to_number(self.bank_column.get())
        desc_column = Google_Service.letter_to_number(self.bank_desc_column.get())
        date_column = Google_Service.letter_to_number(self.bank_date_column.get())

        self.transactions_from_spreadsheet = []
        self.transactions_from_bank = [] # clear prior results

        i = 0
        for oneData in sheet["data"][0]["rowData"]:
            i+=1
            try:
                if "effectiveValue" in oneData["values"][sheet_column] and "numberValue" in oneData["values"][sheet_column]["effectiveValue"]:
                    self.transactions_from_spreadsheet.append(oneData["values"][sheet_column]["effectiveValue"]["numberValue"])
            except IndexError:
                pass
                # this can happen for some inexplicable reason... screw python.

        reader = csv.reader(self.file)
        self.file.seek(0)

        for i in reader:
            length = len(i)

            if length <= bank_column or length <= date_column or length <= desc_column:
                continue

            try:
                row = {'amount':abs(float(i[bank_column])), "pure_amount": i[bank_column]}
                if desc_column:
                    row["description"] = i[desc_column]
                if date_column:
                    row["date"] = i[date_column]

                self.transactions_from_bank.append(row)
            except ValueError:
                continue

        self.result_box.delete(0, "end")  # clear the box before working on it
        disparities, rows_to_highlight \
            = Google_Service.get_disparity_list(self.transactions_from_bank, self.transactions_from_spreadsheet)

        for i in disparities:
            self.result_box.insert("end", Google_Service.formatMoney(i["pure_amount"]) + " " + i["date"] + ' ' + i["description"])

        self.google.generate_new_sheet(sheet_index, rows_to_highlight, sheet_column)



    def run(self):
        self.load_configuration()
        self.initialize_gui()
        self.window.mainloop()

    def update_compute_button(self, event=None):
        if self.ready_to_set:
            self.enter_button["state"] = "active"
        else:
            self.enter_button["state"] = "disabled"

    @property
    def ready_to_set(self):
        return self.is_csv_set and len(self.sheet_box.curselection()) > 0 and len(self.sheet_column.get()) > 0 and len(self.bank_column.get())
