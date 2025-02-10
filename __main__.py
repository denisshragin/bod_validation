# __main__.py

import os
import sys
from datetime import datetime

# ------------------------------------ ## -----------------------------------------
import wx
import openpyxl

# ------------------------------------ ## -----------------------------------------

from helpers import (
                    COLORS,
                    detect_target_columns,
                    calculate_average,
                    color_row,
                    color_bad_final_reading_row
                    )

VERSION = "0.0.1"

# Set the width and the height of the GUI window 
WINDOW_WIDTH = 900
WINDOW_HEIGHT = 700

CONDITIONS = {
            "Initial": 7,
            "Final" : 2,
            "Depletion" : 1,
            "Duplicata" : 30
}

# Set the fonts
header_font = ('Times New Roman', 20)
basic_font = ('Times New Roman', 16)

def main():
    app = wx.App()
    BODProcessorApp(None, title=f"BOD data processor. Version {VERSION}")
    app.MainLoop()

# ------------------------------------ ## ------------------------------------------

# GUI Part

class BODProcessorApp(wx.Frame):
    def __init__(self, parent, title):
        super(BODProcessorApp, self).__init__(parent, title=title, size=(WINDOW_WIDTH, WINDOW_HEIGHT))
        self.InitUI()
        self.Centre()
        self.Show()

        self.file_path = ''

    def InitUI(self):
        panel = wx.Panel(self)
        self.num_selected_files = 0
        
        # Column
        vbox = wx.BoxSizer(wx.VERTICAL)

        select_bod_text = wx.StaticText(panel, label="Select a BOD Instrument:")
        vbox.Add(select_bod_text, 0, wx.EXPAND | wx.ALL, 5)

        # Add a set of radio buttons to choose an instrument
        hbox_radio_bod = wx.BoxSizer(wx.HORIZONTAL)
        self.radio_bod_1 = wx.RadioButton(panel, label='Mario', style=wx.RB_GROUP)
        self.radio_bod_1.SetValue(True)
        hbox_radio_bod.Add(self.radio_bod_1, 0, wx.ALL, 5)
        self.radio_bod_2 = wx.RadioButton(panel, label='Luigi')
        hbox_radio_bod.Add(self.radio_bod_2, 0, wx.ALL, 5)
        self.radio_bod_3 = wx.RadioButton(panel, label='Peach')
        hbox_radio_bod.Add(self.radio_bod_3, 0, wx.ALL, 5)

        vbox.Add(hbox_radio_bod, 0, wx.EXPAND | wx.ALL, 5)

        # --------------------------------------------- ## -----------------------------------

        # Selecting a Directory with the BOD data file to process

        
        # HOW TO SET A FOLDER TO BROWSE FROM
        title_browse_files = wx.StaticText(panel, label="Select a File (.xlsx) to process:")
        title_browse_files.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        vbox.Add(title_browse_files, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        hbox_files_input = wx.BoxSizer(wx.HORIZONTAL)
        self.files_folder_path = wx.TextCtrl(panel, style=wx.TE_PROCESS_ENTER)
        hbox_files_input.Add(self.files_folder_path, 1, wx.EXPAND | wx.ALL, 10)

        btn_browse_files = wx.Button(panel, label="Browse for File")
        btn_browse_files.Bind(wx.EVT_BUTTON, self.OnBrowseFileFolder)
        hbox_files_input.Add(btn_browse_files, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        vbox.Add(hbox_files_input, 0, wx.EXPAND | wx.ALL, 5)


        #  Как сделать чтобы обработать выбранный файл минуя необходимость его выбирать еще один раз
        # A listbox to choose a file to process
        self.list_box_file = wx.ListBox(panel, style=wx.LB_EXTENDED)
        vbox.Add(self.list_box_file, 1, wx.EXPAND | wx.ALL, 10)

        # Add a button to process selected file
        btn_process_file = wx.Button(panel, label="Process Selected File")
        btn_process_file.Bind(wx.EVT_BUTTON, self.OnProcessFile)
        vbox.Add(btn_process_file, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        panel.SetSizer(vbox)


    def OnBrowseFileFolder(self, event):
        # HOW TO DISPLAY ONLY THE .XLSX FILES
        with wx.FileDialog(self, "Open File", wildcard="All files (*.*)|*.*",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dialog:
            if file_dialog.ShowModal() == wx.ID_CANCEL:
                return  # the user changed their mind
            
            # Get the file path from the dialog
            self.file_path = file_dialog.GetPath()
            self.PopulateFileListBox(file_path=self.file_path)
    
    def PopulateFileListBox(self, file_path):
        self.list_box_file.Clear()
        self.list_box_file.Append(self.file_path)

    def OnProcessFile(self, event):
        self.selected_file_path = self.list_box_file.GetStrings()[0]
    
        # Load the workbook and select the active sheet
        wb = openpyxl.load_workbook(self.selected_file_path)
        sheet = wb.active # To fix sheet name if necessary

        # Columns to check (e.g., 'C' for column C)
        target_columns = detect_target_columns(sheet)

        # column_sample_name = 'A'
        # column_to_check_final = 'H'
        # column_to_check_depletion = 'I'
        

        self.previous_sample_number = 'NA'
        self.sample_results = []
        next_sample = True

        # Iterate over the cells in the specified column, skipping the first row
        for row in range(2, sheet.max_row + 1):  # Starts from 2 to skip the first row
            current_sample_number = str(sheet[f'{target_columns["Sample Number"]}{row}'].value).strip()
            if self.previous_sample_number == 'NA' or current_sample_number == self.previous_sample_number:
                color_row(next_sample=next_sample, sheet=sheet, row=row)
            elif current_sample_number != self.previous_sample_number:
                next_sample = not next_sample
                color_row(next_sample=next_sample, sheet=sheet, row=row)
            cell_initial = sheet[f'{target_columns["Init"]}{row}']
            cell_final = sheet[f'{target_columns["Final"]}{row}']
            cell_depletion = sheet[f'{target_columns["Depl"]}{row}']
            
            valid_initial_reading = (isinstance(cell_initial.value, (int, float)) and cell_initial.value > CONDITIONS["Initial"])
            valid_final_reading = (isinstance(cell_final.value, (int, float)) and cell_final.value > CONDITIONS["Final"])
            valid_depletion = (isinstance(cell_depletion.value, (int, float)) and cell_depletion.value > CONDITIONS["Depletion"])

            if not valid_initial_reading:
                cell_initial.fill = COLORS["initial_fill"]
            if not valid_final_reading:
                # cell_final.fill = COLORS["final_fill"]
                color_bad_final_reading_row(sheet=sheet, row=row, color=COLORS["final_fill"])
            if not valid_depletion:
                cell_depletion.fill = COLORS["depletion_fill"]
                # color_bad_final_reading_row(sheet=sheet, row=row, color=COLORS["depletion_fill"])

            if self.previous_sample_number == 'NA':
                self.previous_sample_number = current_sample_number
                if sheet[f'{target_columns["BOD"]}{row}'].value:
                    self.sample_results.append(sheet[f'{target_columns["BOD"]}{row}'].value)

            elif self.previous_sample_number == current_sample_number and valid_final_reading:
                if sheet[f'{target_columns["BOD"]}{row}'].value:
                    self.sample_results.append(sheet[f'{target_columns["BOD"]}{row}'].value)
                if len(self.sample_results) == 2:
                    average, difference = calculate_average(sample_result=self.sample_results)
                    sheet[f'{target_columns["Average"]}{row}'].value = average
                    sheet[f'{target_columns["Difference"]}{row}'].value = difference
                    if difference <= CONDITIONS["Duplicata"]:
                        sheet[f'{target_columns["Average"]}{row}'].fill = COLORS["good_result_fill"]
                        sheet[f'{target_columns["Difference"]}{row}'].fill = COLORS["difference_fill"]
                    self.sample_results.pop(0)
            
            elif self.previous_sample_number != current_sample_number:
                self.sample_results.clear()
                self.previous_sample_number = current_sample_number
                if sheet[f'{target_columns["BOD"]}{row}'].value:
                    self.sample_results.append(sheet[f'{target_columns["BOD"]}{row}'].value)


        wb.save(self.selected_file_path.replace('.', '_colored.'))

# ---------------------------------- ## ----------------------------------------


if __name__ == '__main__':
    main()
