# TODO compile to one single executable file
import logging
import datetime as dt
import tkinter as tk
import tkinter.messagebox
import customtkinter as ctk

from icecream import ic  # V2.1.3

# ic.disable()
ic.configureOutput(includeContext=True)

# from openpyxl import load_workbook
# from openpyxl.styles import Alignment, Font


ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme(
    "blue"
)  # Themes: "blue" (standard), "green", "dark-blue"


# class Window(Frame):
#     def __init__(self, master=None):
#         Frame.__init__(self, master)
#         self.master = master

#         self.pack(fill=BOTH, expand=1)

#         self.startday = StringVar()
#         self.startday.set(dt.datetime.now().strftime("%A"))
#         self.startdate = StringVar()
#         self.startdate.set(dt.datetime.now().strftime("%d.%m.%Y"))
#         self.starttime = StringVar()
#         self.starttime.set(dt.datetime.now().strftime("%H:%M"))

#         # TODO add the other breaks
#         self.breakone = StringVar()

#         self.firstBreakButton = Button(
#             self, text="Fill in first break", command=self.get_current_time
#         )
#         self.firstBreakButton.pack(side=TOP, padx=15, pady=15)

#         automateButton = Button(self, text="Automate", command=self.excel)
#         automateButton.pack(side=BOTTOM, padx=15, pady=15)

#     def get_current_time(self):
#         date = dt.datetime.now().strftime("%H:%M")
#         self.breakone.set(date)

#     def excel(self):
#         # TODO if breaks are empty
#         # -> fill in the end time
#         startday = self.startday.get()
#         startdate = self.startdate.get()
#         starttime = self.starttime.get()
#         breakone = self.breakone.get()
#         self.automate(startday, startdate, starttime, breakone)

#     # TODO close the application when automate button is clicked
#     def automate(self, startday, startdate, starttime, breakone):
#         try:
#             workbook = load_workbook(
#                 filename="C:/Users/POH1SE/main/ata/Homeoffice Korrekturbeleg Arbeitszeiten 4.xlsx"
#             )
#             sheet = workbook.active

#             max_row = 0
#             for row in sheet.iter_rows(values_only=True):
#                 max_row += 1

#             sheet.insert_rows(idx=max_row + 1)

#             # TODO FIXME -> get the correct time for each block
#             # first break block
#             break_end_one = dt.datetime.now()
#             # second working block
#             end_two = dt.datetime.now()
#             # second break block
#             break_end_two = dt.datetime.now()
#             # third working block
#             end_three = dt.datetime.now()

#             # day and date
#             # TODO add translation for weekday
#             weekday = sheet.cell(row=max_row + 1, column=2, value=startday)
#             weekday.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             weekday.font = Font(size=8)

#             date = sheet.cell(row=max_row + 1, column=3, value=startdate)
#             date.alignment = Alignment(horizontal="center", vertical="center")
#             date.font = Font(size=8)

#             # work block 1
#             start_one_cell = sheet.cell(
#                 row=max_row + 1, column=4, value=starttime
#             )
#             start_one_cell.number_format = "h:mm"
#             start_one_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             start_one_cell.font = Font(size=8)

#             end_one_cell = sheet.cell(row=max_row + 1, column=5, value=breakone)
#             end_one_cell.number_format = "h:mm"
#             end_one_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             end_one_cell.font = Font(size=8)

#             # # backup calculation:
#             # # alternative calculating cell values from here instead of using excel formula
#             # first_delta = round((end_one - start_one).total_seconds() / (60 * 60), 2)
#             # sheet.cell(row=max_row + 1, column=6, value=first_delta)

#             # use excel formula to calculate cell values
#             sum_block_one = f'=ROUND(IF(ISBLANK({end_one_cell.coordinate}),"",({end_one_cell.coordinate}-{start_one_cell.coordinate})*24),2)'
#             sum_block_one_cell = sheet.cell(
#                 row=max_row + 1, column=6, value=sum_block_one
#             )
#             sum_block_one_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             sum_block_one_cell.font = Font(size=8)

#             # break block 1
#             formula_break_one = f'=IF(ISBLANK({end_one_cell.coordinate}), "",{end_one_cell.coordinate})'
#             break_one_cell = sheet.cell(
#                 row=max_row + 1, column=7, value=formula_break_one
#             )
#             break_one_cell.number_format = "h:mm"
#             break_one_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             break_one_cell.font = Font(size=8)

#             break_end_one_cell = sheet.cell(
#                 row=max_row + 1, column=8, value=break_end_one.strftime("%H:%M")
#             )
#             break_end_one_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             break_end_one_cell.font = Font(size=8)

#             # work block 2
#             formula_start_two = f'=IF(ISBLANK({break_end_one_cell.coordinate}), "",{break_end_one_cell.coordinate})'
#             start_two_cell = sheet.cell(
#                 row=max_row + 1, column=9, value=formula_start_two
#             )
#             start_two_cell.number_format = "h:mm"
#             start_two_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             start_two_cell.font = Font(size=8)

#             end_two_cell = sheet.cell(
#                 row=max_row + 1, column=10, value=end_two.strftime("%H:%M")
#             )
#             end_two_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             end_two_cell.font = Font(size=8)

#             sum_block_two = f'=ROUND(IF(ISBLANK({start_two_cell.coordinate}),"",({end_two_cell.coordinate}-{start_two_cell.coordinate})*24),2)'
#             sum_block_two_cell = sheet.cell(
#                 row=max_row + 1, column=11, value=sum_block_two
#             )
#             sum_block_two_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             sum_block_two_cell.font = Font(size=8)

#             # break block 2
#             formula_break_two = f'=IF(ISBLANK({end_two_cell.coordinate}), "",{end_two_cell.coordinate})'
#             break_two_cell = sheet.cell(
#                 row=max_row + 1, column=12, value=formula_break_two
#             )
#             break_two_cell.number_format = "h:mm"
#             break_two_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             break_two_cell.font = Font(size=8)

#             break_end_two_cell = sheet.cell(
#                 row=max_row + 1,
#                 column=13,
#                 value=break_end_two.strftime("%H:%M"),
#             )
#             break_end_two_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             break_end_two_cell.font = Font(size=8)

#             # work block 3
#             formula_start_three = f'=IF(ISBLANK({break_end_two_cell.coordinate}), "",{break_end_two_cell.coordinate})'
#             start_three_cell = sheet.cell(
#                 row=max_row + 1, column=14, value=formula_start_three
#             )
#             start_three_cell.number_format = "h:mm"
#             start_three_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             start_three_cell.font = Font(size=8)

#             end_three_cell = sheet.cell(
#                 row=max_row + 1, column=15, value=end_three.strftime("%H:%M")
#             )
#             end_three_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             end_three_cell.font = Font(size=8)

#             sum_block_three = f'=ROUND(IF(ISBLANK({start_three_cell.coordinate}),"",({end_three_cell.coordinate}-{start_three_cell.coordinate})*24),2)'
#             sum_block_three_cell = sheet.cell(
#                 row=max_row + 1, column=16, value=sum_block_three
#             )
#             sum_block_three_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             sum_block_three_cell.font = Font(size=8)

#             # sum work and sum break
#             formula_sum_work = f'=IF(ISBLANK({date.coordinate}),"", {sum_block_one_cell.coordinate}+{sum_block_two_cell.coordinate}+{sum_block_three_cell.coordinate})'
#             formula_sum_work_cell = sheet.cell(
#                 row=max_row + 1, column=17, value=formula_sum_work
#             )
#             formula_sum_work_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             formula_sum_work_cell.font = Font(size=8)

#             formula_sum_break = f'=IF(ISBLANK({date.coordinate}),"", IF(ISBLANK({break_one_cell.coordinate}),0,({break_end_one_cell.coordinate}-{break_one_cell.coordinate})*24)+IF(ISBLANK({break_two_cell.coordinate}),0,({break_end_two_cell.coordinate}-{break_two_cell.coordinate})*24))'
#             formula_sum_break_cell = sheet.cell(
#                 row=max_row + 1, column=18, value=formula_sum_break
#             )
#             formula_sum_break_cell.alignment = Alignment(
#                 horizontal="center", vertical="center"
#             )
#             formula_sum_break_cell.font = Font(size=8)

#             workbook.save(
#                 filename="C:/Users/POH1SE/main/ata/Homeoffice Korrekturbeleg Arbeitszeiten 4.xlsx"
#             )

#         except Exception as err:
#             print(err)
#             logging.error("%s", err, exc_info=True)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.startday = tk.StringVar()
        self.startday.set(dt.datetime.now().strftime("%A"))
        self.startdate = tk.StringVar()
        self.startdate.set(dt.datetime.now().strftime("%d.%m.%Y"))
        self.starttime = tk.StringVar()
        self.starttime.set(dt.datetime.now().strftime("%H:%M"))

        # TODO add the other breaks
        self.breakone = tk.StringVar()

        # configure window
        self.title("ATA")
        self.geometry(f"{800}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = ctk.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="Homeoffice\nAutomation",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.appearance_mode_label = ctk.CTkLabel(
            self.sidebar_frame, text="Appearance Mode:", anchor="w"
        )
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=["Light", "Dark", "System"],
            command=self.change_appearance_mode_event,
        )
        self.appearance_mode_optionemenu.grid(
            row=6, column=0, padx=20, pady=(10, 10)
        )

        self.scaling_label = ctk.CTkLabel(
            self.sidebar_frame, text="UI Scaling:", anchor="w"
        )
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = ctk.CTkOptionMenu(
            self.sidebar_frame,
            values=["80%", "90%", "100%", "110%", "120%"],
            command=self.change_scaling_event,
        )
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # create main entry and button
        self.entry = ctk.CTkEntry(self, placeholder_text="Enter data here")
        self.entry.grid(
            row=3,
            column=1,
            columnspan=1,
            padx=(20, 0),
            pady=(20, 20),
            sticky="nsew",
        )

        self.main_button_1 = ctk.CTkButton(
            master=self,
            fg_color="transparent",
            border_width=2,
            text_color=("gray10", "#DCE4EE"),
            command=self.button_event,
        )
        self.main_button_1.grid(
            row=3, column=2, padx=(20, 20), pady=(20, 20), sticky="nsew"
        )

        self.string_input_button = ctk.CTkButton(
            self,
            text="Open CTkInputDialog",
            command=self.open_input_dialog_event,
        )
        self.string_input_button.grid(row=2, column=2, padx=20, pady=(10, 10))

        self.firstBreakButton = ctk.CTkButton(
            self, text="Fill in first break", command=self.get_current_time
        )
        self.firstBreakButton.grid(row=2, column=1, padx=20, pady=20)

        # output label
        self.output_label = ctk.CTkLabel(master=self, text="")
        self.output_label.grid(row=1, column=1, padx=20, pady=20, sticky="nsew")

        # create slider and progressbar frame
        self.slider_progressbar_frame = ctk.CTkFrame(
            self, fg_color="transparent"
        )
        self.slider_progressbar_frame.grid(
            row=0,
            column=1,
            columnspan=2,
            padx=(20, 20),
            pady=(20, 20),
            sticky="nsew",
        )
        self.slider_progressbar_frame.grid_columnconfigure(0, weight=1)
        self.progressbar_1 = ctk.CTkProgressBar(self.slider_progressbar_frame)
        self.progressbar_1.grid(
            row=1, column=0, padx=(20, 10), pady=(10, 10), sticky="ew"
        )

        # set default values
        self.appearance_mode_optionemenu.set("System")
        self.scaling_optionemenu.set("100%")
        self.progressbar_1.configure(mode="indeterminnate")
        self.progressbar_1.start()

    def open_input_dialog_event(self):
        dialog = ctk.CTkInputDialog(
            text="Type in a number:", title="CTkInputDialog"
        )
        print("CTkInputDialog:", dialog.get_input())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)

    def button_event(self):
        print("Click")

    def get_current_time(self):
        date = dt.datetime.now().strftime("%H:%M")
        self.breakone.set(date)
        current_text = self.output_label.cget("text")
        new_text = current_text + date + "\n"
        self.output_label.configure(text=new_text)


if __name__ == "__main__":
    logging.basicConfig(
        filename="ata.log",
        level=logging.DEBUG,
        format="[%(levelname)s]::%(name)s::(%(asctime)s) - %(message)s",
    )

    app = App()
    app.mainloop()
