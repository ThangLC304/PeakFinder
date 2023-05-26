import tkinter as tk
from tkinter import ttk
import customtkinter 
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import os
from pathlib import Path
from tqdm import tqdm
import logging
import threading
import time
import numpy as np
import openpyxl
import json
import shutil

from Libs.utils import *
from Libs.findpeaks import *

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s:%(name)s:%(message)s')

file_handler = logging.FileHandler('app.log')
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)

# Create core folders
root_folder = Path(__file__).parent
init_core_folders(root_folder)

class Parameters():
    def __init__(self, master, row, text, default_value):

        # create a label 
        self.lbl = customtkinter.CTkLabel(master, text=text)
        self.lbl.grid(row=row, column=0, padx=20, pady=20, sticky='nswe')

        # create a entry
        self.entry = customtkinter.CTkEntry(master)
        self.entry.grid(row=row, column=1, padx=20, pady=20, sticky='nswe')
        self.entry.insert(0, default_value)

    def get(self):
        return self.entry.get()
    
    def set(self, value):
        self.entry.delete(0, tk.END)
        self.entry.insert(0, value)


class History():    
    def __init__(self):
        self.history = self.load()

    def load(self):
        try:
            with open("Log/history.json", "r") as f:
                history = json.load(f)
        except:
            history = {}
            self.save(history)
        return history
    
    def save(self, history=None):
        if history is None:
            history = self.history

        with open("Log/history.json", "w") as f:
            json.dump(history, f, indent=4)

    def add(self, data_type, key):
        if data_type not in self.history:
            self.history[data_type] = {}
        if key not in self.history[data_type]:
            self.history[data_type][key] = 1
        else:
            self.history[data_type][key] += 1
        self.save()

    def most_selected(self, data_type, key_list):
        if data_type not in self.history:
            return key_list[0]
        else:
            # return the key of [data_type] that is also in key_list, with highest value
            for key in sorted(self.history[data_type], key=self.history[data_type].get, reverse=True):
                if key in key_list:
                    return key
            return key_list[0]


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        # setup geometry
        window_width = 500
        window_height = 800
        self.geometry(f"{window_width}x{window_height}+{int(self.winfo_screenwidth()/2 - window_width/2)}+{int(self.winfo_screenheight()/2 - window_height/2)}")
        self.title("File Manager")

        # Define your main frame and set its size
        self.MAIN = customtkinter.CTkFrame(self, width=window_width, height=window_height)
        # center it
        self.MAIN.place(anchor='c', relx=.5, rely=.5)

        self.name = None
        self.output_dir = "Output"
        self.history = History()
        self.mode = "individual"

        button_config = {"font": ('Helvetica', 16), "width": 150, "height": 40}

        # Widget 1: Button to load excel file
        self.btn_load = customtkinter.CTkButton(self.MAIN, 
                                                text="Load Excel File", 
                                                command=self.load_excel_file)
        self.btn_load.configure(**button_config)
        self.btn_load.grid(row=0, columnspan=2, padx=20, pady=20, sticky='nswe')


        # Create widgets for input parameters:
        tolerance = 0
        minPeakDistance = 0
        minMaximaValue = np.nan
        maxMaximaValue = np.nan

        self.ToleranceEntry = Parameters(self.MAIN, 1, "Tolerance", tolerance)
        self.minPeakDistance = Parameters(self.MAIN, 2, "minPeakDistance", minPeakDistance)
        self.minMaximaValue = Parameters(self.MAIN, 3, "minMaximaValue", minMaximaValue)
        self.maxMaximaValue = Parameters(self.MAIN, 4, "maxMaximaValue", maxMaximaValue)

        self.excludeOnEdges = customtkinter.CTkCheckBox(self.MAIN, text="excludeOnEdges")
        self.excludeOnEdges.grid(row=5, column=0, padx=20, pady=20, sticky='nswe')

        # Widget 2: A drop down list to select the columns inside the excel file, only visible after loading the excel file
        self.ColumnLabel = customtkinter.CTkLabel(self.MAIN, 
                                                  text="Select Column")
        self.ColumnLabel.grid(row=6, column=0, padx=20, pady=20, sticky='nswe')

        self.ColumnOption = customtkinter.CTkOptionMenu(self.MAIN, state="readonly")
        self.ColumnOption.grid(row=6, column=1, padx=20, pady=20, sticky='nswe')


        # Widget 3: A drop down list to select the sheets inside the excel file, only visible after loading the excel file
        self.SheetLabel = customtkinter.CTkLabel(self.MAIN,
                                                 text="Select Sheet")

        self.SheetOption = customtkinter.CTkOptionMenu(self.MAIN, state="readonly")
        self.SheetOption.grid(row=7, column=1, padx=20, pady=20, sticky='nswe')

        # Widget 4: A button to Find Peaks, only visible after loading the excel file
        self.btn_find_peaks = customtkinter.CTkButton(self.MAIN, 
                                                      text="Find Peaks", 
                                                      command=self.find_peaks_thread)

        # Widget 5: A button to save the peaks to a file, only visible after finding the peaks
        self.btn_save_peaks = customtkinter.CTkButton(self.MAIN, 
                                                      text="Save Peaks", 
                                                      command=self.save_peaks)
        
        # Widget 6: A button to draw the peaks, only visible after finding the peaks
        self.btn_draw_peaks = customtkinter.CTkButton(self.MAIN,
                                                      text="Draw Peaks",
                                                      command=self.draw_peaks)

        # Bind both dropdown list selection to preprocess, every time an option is selected, the preprocess function will be called
        self.ColumnOption.configure(command=self.preprocess)
        self.SheetOption.configure(command=self.preprocess)


    def create_progress_window(self, title="Analysis Progress", text="Analyzing..."):
        progress_window = tk.Toplevel(self)
        progress_window.title(title)
        progress_window.geometry("300x100")

        progress_label = tk.Label(progress_window, text=text, font=('Helvetica', 12))
        progress_label.pack(pady=(10, 0))

        progress_bar = ttk.Progressbar(progress_window, mode='determinate', length=200)
        progress_bar.pack(pady=(10, 20))

        return progress_bar


    def preprocess(self, value = None, sheet_name = None):
        print(f"Column selected: {self.ColumnOption.get()}")
        print(f"Sheet selected: {self.SheetOption.get()}")

        logger.info(f"Column selected: {self.ColumnOption.get()}")
        logger.info(f"Sheet selected: {self.SheetOption.get()}")

        if sheet_name is None:
            self.sheet_name = self.SheetOption.get()
            print(f"Sheet_name was None, load sheet name from Option: {self.sheet_name}")
        else:
            self.sheet_name = sheet_name
            print(f"Loaded sheet name as given: {self.sheet_name}")

        print(f"Actual loaded sheet name: {self.sheet_name}")
        logger.info(f"Actual loaded sheet name: {self.sheet_name}")
        
        if self.sheet_name == "All Sheets":
            self.mode = "batch"
            return
        else:
            self.mode = "individual"

        print(f"Mode: {self.mode}")
        logger.info(f"Mode: {self.mode}")

        column_name = self.ColumnOption.get()

        # Load data according to the sheet & column selected
        try:
            self.df = read_clean_excel(self.excel_path, sheet_name=self.sheet_name)
        except Exception as e:
            logger.error(e)
            # use tkinter to show message box
            tk.messagebox.showerror("Error", "Excel file was not loaded!")
            return

        self.yvalues = self.df[column_name].values
        if len(self.yvalues) > 0:
            self.yvalues = np.array(self.yvalues, dtype=np.float64)
            logger.info(f"Loaded data ! Number of data points: {len(self.yvalues)}")
        else:
            logger.info("Selected column is empty !")
        
        # set value of self.tolerance to
        tolerance = np.std(self.yvalues)
        self.ToleranceEntry.set(tolerance)

    def load_excel_file(self):

        self.excel_path = tk.filedialog.askopenfilename(initialdir=os.getcwd(), title="Select Excel File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        if self.excel_path:
            logger.info(f"Loading {self.excel_path}")
            self.name = Path(self.excel_path).name
            self.df = read_clean_excel(self.excel_path)
            self.columns = self.df.columns
            self.output_path = os.path.join(self.output_dir, self.name)
            logger.info(f"Columns in the excel file: {self.columns}")

            try:
                os.startfile(self.excel_path)
            except:
                pass

            self.ColumnOption.configure(values=self.columns)
            column_init = self.history.most_selected("column", self.columns)
            self.ColumnOption.set(column_init)

            wb = openpyxl.load_workbook(self.excel_path)
            self.sheet_names = wb.sheetnames
            self.SheetOption.configure(values=["All Sheets"] + self.sheet_names)
            self.sheet_name = self.sheet_names[0]
            self.SheetOption.set(self.sheet_name)

            self.preprocess()

            self.btn_find_peaks.grid(row=9, columnspan=2, padx=20, pady=20, sticky='nswe')
            # set value of self.tolerance to the value

    def find_peaks_thread(self):
        print("Start finding peaks in a new thread !")
        logger.info("Start finding peaks in a new thread !")

        if self.mode == "individual":
            print("Individual mode !")
            logger.info("Individual mode !")
            thread = threading.Thread(target=self.find_peaks)
            thread.start()
        elif self.mode == "batch":
            print("Batch mode !")
            logger.info("Batch mode !")
            batch_progress_bar = self.create_progress_window()

            for i, sheet_name in enumerate(self.sheet_names):
                batch_progress_bar['value'] = (i+1)/len(self.sheet_names) * 100
                batch_progress_bar.update()

                self.sheet_name = sheet_name

                thread = threading.Thread(target=self.find_and_save(sheet_name))
                thread.start()

            batch_progress_bar.master.destroy()
            # Display a messagebox
            tk.messagebox.showinfo("Analysis completed", f"Analyzed file is saved in {self.output_dir}")
                
        else:
            logger.error("Mode not recognized !")

    def find_and_save(self, sheet_name):
        self.preprocess(sheet_name = sheet_name)
        self.find_peaks()
        self.save_peaks()
        self.draw_peaks(mode="save")
        logger.info(f"Finished finding & saving peaks of {sheet_name} !")

    def find_peaks(self):

        # self.history.add("sheet", sheet_name)
        column_name = self.ColumnOption.get()
        self.history.add("column", column_name)

        tolerance = float(self.ToleranceEntry.get())
        minPeakDistance = float(self.minPeakDistance.get())
        minMaximaValue = np.nan if self.minMaximaValue.get() == "NaN" or self.minMaximaValue.get() == "" else float(self.minMaximaValue.get())
        maxMaximaValue = np.nan if self.maxMaximaValue.get() == "NaN" or self.maxMaximaValue.get() == "" else float(self.maxMaximaValue.get())
        excludeOnEdges = self.excludeOnEdges.get()

        logger.info(f"tolerance: {tolerance}")
        logger.info(f"minPeakDistance: {minPeakDistance}")
        logger.info(f"minMaximaValue: {minMaximaValue}")
        logger.info(f"maxMaximaValue: {maxMaximaValue}")
        logger.info(f"excludeOnEdges: {excludeOnEdges}")


        progress_bar = self.create_progress_window(title=self.sheet_name, text="Finding peaks ...")

        # find peaks
        xvalues, yvalues, maxima, minima = PeakFinder(progress_bar,
                                                      self.yvalues, 
                                                      tolerance,
                                                      minPeakDistance, 
                                                      minMaximaValue, 
                                                      maxMaximaValue, 
                                                      excludeOnEdges)

        self.xvalues = xvalues
        self.yvalues = yvalues
        self.maxima = maxima
        self.minima = minima

        if len(self.maxima) > 0 and len(self.minima) > 0:
            logger.info(f"Found {len(self.maxima)} maxima and {len(self.minima)} minima")
            self.btn_save_peaks.grid(row=10, column=0, padx=20, pady=20, sticky='nswe')
            self.btn_draw_peaks.grid(row=10, column=1, padx=20, pady=20, sticky='nswe')
        else:
            logger.info("No peaks found")
            return
        
        # close the progress bar
        progress_bar.master.destroy()
        
    def save_peaks(self):
        if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)

        if self.name != None:
            

            # make a dataframe to store the peaks
            df_maxima, df_minima = make_df(self.xvalues, self.yvalues, self.maxima, self.minima)
            
            if len(df_maxima) > 0 and len(df_minima) > 0:
                logger.info(f"Found {len(df_maxima)} maxima and {len(df_minima)} minima")
            else:
                print("No peaks found")
                logger.info("No peaks found")
                return
            
            # save the dataframe to 2 sheets named maxima and minima of an excel file
            print("Saving peaks ...")
            output_path = self.output_path
            logger.info(f"Saving peaks at {output_path}, sheet name: {self.sheet_name}...")

            # Copy the original excel file to the output directory
            if not os.path.exists(output_path):
                shutil.copy(self.excel_path, output_path)
                logger.info(f"First time analysis, cloned {self.name} to {self.output_dir}")
            else:
                # open self.sheet_name
                temp_df = read_clean_excel(output_path, sheet_name=self.sheet_name)
                # if the all "xMaxima", "yMaxima", "xMinima", "yMinima" columns exist, skip
                temp_count = 0
                for temp_column in temp_df.columns:
                    if "maxima" or "minima" in temp_column.lower():
                        temp_count += 1
                if temp_count >= 4:
                    logger.info(f"Found maxima and minima columns in {self.sheet_name}, skip saving")
                    return

            try:
                append_df_to_excel(output_path, df_maxima, sheet_name=self.sheet_name, index=False, startrow=2, startcol=len(self.columns))
                logger.info(f"Saved maxima to sheet {self.sheet_name} of {output_path}")
            except Exception as e:
                print(e)
                logger.info(f"Failed to save maxima to sheet {self.sheet_name} of {output_path}")
                logger.info(e)

            try:
                append_df_to_excel(output_path, df_minima, sheet_name=self.sheet_name, index=False, startrow=2, startcol=len(self.columns)+2)
                logger.info(f"Saved minima to sheet {self.sheet_name} of {output_path}")
            # except as e, print e
            except Exception as e:
                print(e)
                logger.info(f"Failed to save minima to sheet {self.sheet_name} of {output_path}")
                logger.info(e)


    def draw_peaks(self, mode="display"):

        assert mode in ["display", "save"]

        figure = draw_plot(self.xvalues, self.yvalues, self.maxima, self.minima)

        if mode=="display":
            # Make a top level window
            self.top = tk.Toplevel(self)
            self.top.title("Peaks")
            self.top.geometry("800x400")

            # Make a canvas
            self.canvas = FigureCanvasTkAgg(figure, master=self.top)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

            # make a button on top right corner to save the figure
            self.btn_save_figure = tk.Button(self.top, text="Save", command=lambda: self.save_pictures(figure))
            self.btn_save_figure.pack(side=tk.RIGHT)

        if mode=="save":
            self.save_pictures(figure)

    def save_pictures(self, figure):
        dir_name = self.name.split(".")[0]
        image_name = self.name.split(".")[0] + "_" + self.sheet_name + ".png"
        output_path = os.path.join(self.output_dir, dir_name, image_name)
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        if os.path.exists(output_path):
            logger.info(f"{output_path} already exists, skip saving")
            return
        figure.savefig(output_path, dpi=300)
        logger.info(f"Saved peaks plot to {output_path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
