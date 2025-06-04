"""
Functional goals of this program:
    1) To make it easy to put both pictures and descriptions of every crystal we find immediately into Box without having to open up Box and navigate through a complex Excel sheet
    that noramlly requires very repetitive data entry (and in which we often miss things).
        - This will be accomplished by using the Box SDK with Python to access Box and by creating a Tkinter GUI that is easy to interact with. The GUI will contain reference fields to 
        be filled out and buttons for operating the microscope in an integrated fashion. It may also contain other operations as described in Goal 2.
    2) To incorporate the information from all the previous experimental steps into a single place so that we can track easily (and with less tedium) how our experiments are going.
        - This will be accomplished by creating a background database that includes all of the crystal conditions that we normally use, as well as GUI-led steps for uploading crystal
        optimization conditions (which are often difficult to keep track of on paper). If possible and deemed necessary, this GUI may lead the user to operate a web-sourced crystal 
        optimization condition generator (https://hamptonresearch.com/make-tray.php) and then will immediately scrape the data into the database to be referenced later on in the 
        Excel-sheet editing steps.
        https://docs.python.org/3/library/tkinter.html
"""

#https://byu.app.box.com/developers/console

#Imports
#General imports
import pandas as pd
import os
import shutil
import json
from openpyxl import Workbook, load_workbook
from datetime import datetime
import time

#GUI imports
#https://tkdocs.com/tutorial/intro.html#audience
#https://tkdocs.com/tutorial/firstexample.html#design
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog

#Box integration imports:
#https://github.com/box/box-python-sdk-gen/tree/main
import box_sdk_gen
from box_sdk_gen import BoxClient, BoxDeveloperTokenAuth

#SebaView integration imports:
#https://codezup.com/automate-windows-tasks-with-python-win32-library/
#https://pywinauto.readthedocs.io/en/latest/getting_started.html
import psutil
import pywinauto
from pywinauto.application import Application
from pywinauto import timings
import pyautogui
from pynput import mouse
import pywinauto.keyboard

#Packaging stuff:
#https://realpython.com/pyinstaller-python/

#Paths
script_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(script_dir,"crystaldex_icon.png")

home = os.path.expanduser("~")
downloads = os.path.join(home, "Downloads")
crystal_pictures = os.path.join(script_dir,"Crystal_Pictures")
#Fix this so it only moves .jpegs and move it to the right area.

def on_click(x,y,button,pressed):
                global mouse_is_down
                mouse_is_down = pressed

listener = mouse.Listener(on_click=on_click)
listener.start()

class CrystalDex_main:
    def __init__(self):
        root=Tk()
        self.root = root
        self.root.title("CrystalDex")
        icon = PhotoImage(file=icon_path)
        self.root.iconphoto(True,icon)
        self.root.minsize(700,600)
        #Make the window resizable:
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)
        self.crystallization_chaperone_values = ["1TEL","2TEL","3TEL","4TEL","5TEL","6TEL"]
        self.crystal_screen_values = ["Crystal_Screen","Index","PEG_Custom","PEG_Ion","Salt_Rx","Wizard"]
        token = 'BSMuw9I5ElwQ60CF7cMNmYxsRDGMyvqN'
        auth: BoxDeveloperTokenAuth = BoxDeveloperTokenAuth(token=token)
        self.client: BoxClient = BoxClient(auth=auth)

    def load_SeBaView(self):
        exe_path = None
        if os.path.exists("SeBaView_path_file.json"):
            with open("SeBaView_path_file.json", "r") as s:
                exe_path = json.load(s).get("SeBaView_path")
        if not exe_path or not os.path.exists(exe_path):
            # Ask user to locate it if not found or invalid
            exe_path = filedialog.askopenfilename(
                title="Select the SeBaView executable",
                filetypes=[("Executable files", "*.exe")]
            )
        if exe_path:
            #First, kill any current SeBaView window, then proceed.
            for proc in psutil.process_iter(attrs=["pid", "name", "exe"]):
                try:
                    if proc.info["name"] and "SeBaView" in proc.info["name"]:
                        print(f"Killing SeBaView process: PID {proc.info['pid']}")
                        proc.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            time.sleep(2)

            with open("SeBaView_path_file.json", "w") as s:
                json.dump({"SeBaView_path": exe_path}, s)
            SeBaView = Application(backend="uia").start(exe_path)
            time.sleep(10)
            for i, w in enumerate(SeBaView.windows()):
                print(f'[{i}] Title: {w.window_text()}')
            try:
                # Try to wait for the first active window
                SeBaView_main_window = SeBaView.window(title_re=".*SeBaView.*")
                #timings.wait_until_passes(15, 1, lambda: main_window.exists() and main_window.is_visible())
                SeBaView_wrapper = SeBaView_main_window.wrapper_object()
                SeBaView_wrapper.set_focus()
                print("SeBaView main window focused successfully.")
            except Exception as e:
                print("Failed to find or focus the SeBaView window.")
                print(f"Error: {e}")

            """
            #This code is really buggy and needs to be moved to another section, but it's useful for figuring out the location of buttons in SeBaView.
            SeBaView_wrapper_rect = SeBaView_main_window.rectangle()
            try:
                while True:
                    if mouse_is_down:
                        x, y = pyautogui.position()
                        rel_x = x - SeBaView_wrapper_rect.left
                        rel_y = y - SeBaView_wrapper_rect.top
                        print(f'Mouse pressed relative to window: ({rel_x},{rel_y})')
                    time.sleep(0.1)
            except KeyboardInterrupt:
                pass
            """
            for i in range(2):
                SeBaView_wrapper.click_input(coords=(60, 165))  #This accesses the camera connecting button.
            return SeBaView_wrapper

    def identify_subwell(self,SeBaView_wrapper,date_set_var,crystal_screen_var,target_protein_var,target_protein_top_left_stock_concentration_var,target_protein_top_right_stock_concentration_var,target_protein_bottom_left_stock_concentration_var,crystallization_chaperone_var,custom_tags_values):
        self.clear_widgets()
        self.add_menu()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_width // 3}x{screen_height}+0+0")
        subwell_frame = ttk.Frame(self.root,padding="3 3 12 12")
        subwell_frame.grid(column=0,row=0,sticky=(N,W))
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)

        well_column_label = ttk.Label(subwell_frame,text="Well column:")
        well_column_label.grid(column=1,row=1)
        well_column_values = ['A','B','C','D','E','F','G','H']
        well_column_var = StringVar()
        well_column_drop_down = ttk.Combobox(subwell_frame,textvariable=well_column_var,values=well_column_values)
        well_column_drop_down.grid(column=2,row=1)

        well_row_label = ttk.Label(subwell_frame,text="Well row:")
        well_row_label.grid(column=1,row=2)
        well_row_values = [1,2,3,4,5,6,7,8,9,10,11,12]
        well_row_var = IntVar()
        well_row_drop_down = ttk.Combobox(subwell_frame,textvariable=well_row_var,values=well_row_values)
        well_row_drop_down.grid(column=2,row=2)

        subwell_values = ['top_left','top_right','bottom_left']
        subwell_label = ttk.Label(subwell_frame,text="subwell:")
        subwell_label.grid(column=1,row=5,sticky=(N,W))
        subwell_var = StringVar()
        subwell_drop_down = ttk.Combobox(subwell_frame,textvariable=subwell_var,values=subwell_values)
        subwell_drop_down.grid(column=2,row=5)

        possible_salt_crystals_label = ttk.Label(subwell_frame,text="Possibly a salt crystal")
        possible_salt_crystals_label.grid(column=1,row=6)
        possible_salt_crystals_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=possible_salt_crystals_var).grid(column=2,row=6)

        precipitation_label = ttk.Label(subwell_frame,text="Precipitation present")
        precipitation_label.grid(column=1,row=7)
        precipitation_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=precipitation_var).grid(column=2,row=7)

        ttk.Button(subwell_frame,text="Take and Save Picture",
                command=lambda: self.take_picture(
                     SeBaView_wrapper,
                     f'{crystallization_chaperone_var}_{target_protein_var}_{crystal_screen_var}_{well_column_var.get()}{well_row_var.get()}_{subwell_var.get()}'
                )).grid(column=1,row=14,sticky=(N,W))
        
        for child in subwell_frame.winfo_children():
            child.grid_configure(padx=15,pady=15)

        self.root.deiconify()      # Restore the window if minimized
        self.root.lift()
        self.root.focus_force()

        #Eventually I'll move this code
        for filename in os.listdir(downloads):
            file_path = os.path.join(downloads, filename)
            if os.path.isfile(file_path) and filename.lower().endswith(('.jpeg','.jpg')) and os.path.getmtime(file_path)<1:
                try:
                    shutil.move(file_path, crystal_pictures)
                    print(f"Moved: {filename}")
                except Exception as e:
                    print(f"Failed to move {filename}: {e}")
        
    def take_picture(self,SeBaView_wrapper,image_title):
        SeBaView_wrapper.set_focus() #Is this needed? Test later...
        SeBaView_wrapper.click_input(coords=(55, 70))  #This accesses the save as button.
        time.sleep(3)
        pywinauto.keyboard.send_keys(f"{image_title}{{ENTER}}") #I believe this code will let me type into whatever location is currently focused, but I haven't tested it yet.

    def add_menu(self):
        menu = Menu(self.root)
        menu.add_command(label='Home',command=self.startup)
        menu.add_command(label="Help",command=self.Help)
        self.root.config(menu=menu)

    def clear_widgets(self):
        for widget in self.root.winfo_children():
            if isinstance(widget,ttk.Frame):
                widget.destroy()

    def Help(self):
        self.clear_widgets()
        self.add_menu()
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)
        helpframe = ttk.Frame(self.root,padding='3 3 12 12')
        helpframe.grid(column=0,row=0,sticky=(N,W,E,S))
        ttk.Label(helpframe,text="Welcome to CrystalDex, your helper for recording data from protein crystallization experiments!").grid(column=0,row=0,sticky=(N,E,W))
        helptext = "This program functions by accessing Box and syncing with Excel sheets that contain links to every picture you take.\nCrystalDex allows you to run the microscope application within its GUI and prompts you to measure and label each crystal.\nIt then synchronizes all the crystallization screen data from its library of screens with each crystal picture taken.\nThere are other subprograms in this app that allow you to upload new crystallization screens into its library (such as for optimization screens). \nFor more assistance, reach out to miles.j.bradford@outlook.com"
        ttk.Label(helpframe,text=helptext).grid(column=0,row=1,sticky=(N,E,W))

    def Index_Tray(self,date_set_var,crystal_screen_var,target_protein_var,target_protein_top_left_stock_concentration_var,target_protein_top_right_stock_concentration_var,target_protein_bottom_left_stock_concentration_var,crystallization_chaperone_var,custom_tags_values):
        indexable = False
        file_id = '1861370891462'
        file_download = self.client.downloads.download_file(file_id).read()
        with open("Crystal_Trays_Library.xlsx","wb") as c:
            c.write(file_download) #this writes (or overwrites) a file into the working computer with the download data from Box.
        print("Saving to:", os.path.abspath("Crystal_Trays_Library.xlsx"))
        #df = pd.read_excel("download.xlsx")
        #print(df)
        wb = load_workbook(filename=os.path.abspath("Crystal_Trays_Library.xlsx"))
        date = None
        crystal_screen = None
        target_protein = None

        try:
            date = datetime.strptime(date_set_var,"%m.%d.%Y")
            print(f'indexing! date: {date_set_var}')
            date = date.strftime('%m-%d-%Y')
            indexable = True
        except ValueError:
            messagebox.showerror(title="Date Error",message="You attempted to put in an invalid date. Please use the style: 01.01.2025")
        if crystal_screen_var in self.crystal_screen_values:
            crystal_screen = crystal_screen_var
            indexable = True
        else: 
            messagebox.showerror(title="Crystal Screen Does Not Exist",message="The crystal screen you attempted to reference does not exist.")
            indexable = False
        if target_protein_var is not None:
            target_protein = target_protein_var
            indexable = True
        else:
            messagebox.showerror(title="No Protein Target",message="You neglected to enter a protein target. (CrystalDex can't index nothingness!)")
            indexable = False
        custom_tags_list = [tag.strip() for tag in custom_tags_values.split(',') if tag.strip()]
        if indexable:
            ws_possible_duplicate_count = 0
            for ws in wb:
                if all(term in ws.title for term in [date,crystal_screen,target_protein]):
                    ws_possible_duplicate_count += 1
            if ws_possible_duplicate_count >0:
                print(f'At least one previously indexed tray was found that shares a date, screen, and target protein with the current tray. \nPlease review the following to ensure no duplicate trays are indexed!')#change this to a Tkinter frame.
                for ws in wb:
                    if all(term in ws.title for term in [date,crystal_screen,target_protein]):
                        print(ws['K1'])
                print(f'If none of the above match your tray, click here.')#Change this to a tkinter button...
            if ws_possible_duplicate_count == 0:
                print(f"No trays found with these stats; generating new tray!")
                new_worksheet = wb.copy_worksheet(wb["Mastercopy"])
                full_title = f'{target_protein}_{crystal_screen}_{date}_1'
                short_title = full_title[:26]
                print(f'short_title: {short_title}')
                new_worksheet.title = short_title
                all_tags = [date_set_var,crystal_screen_var,target_protein_var,target_protein_top_left_stock_concentration_var,target_protein_top_right_stock_concentration_var,target_protein_bottom_left_stock_concentration_var,crystallization_chaperone_var,custom_tags_values]
                new_worksheet['K1'] = ','.join(map(str,all_tags))
                new_worksheet['D1'] = str(date_set_var)
                new_worksheet['D2'] = str(crystallization_chaperone_var)
                new_worksheet['D3'] = str(crystal_screen_var)
                new_worksheet['D4'] = str(target_protein_var)
                new_worksheet['D5'] = str(custom_tags_values)
                new_worksheet['H1'] = str(target_protein_top_left_stock_concentration_var)
                new_worksheet['H2'] = str(target_protein_top_right_stock_concentration_var)
                new_worksheet['H3'] = str(target_protein_bottom_left_stock_concentration_var)
                wb.save(filename=os.path.abspath("Crystal_Trays_Library.xlsx"))
                """
                Re-enable these lines once ready to attach to Box.
                self.client.uploads.upload_file_version(
                    attributes=box_sdk_gen.UploadFileAttributesParentField(name="Crystal_Trays_Library.xlsx",
                        id="1862599427539"),
                    file_id="1862599427539",
                    file=open(os.path.abspath("Crystal_Trays_Library.xlsx"),"rb")
                )
                """
            SeBaView_wrapper = self.load_SeBaView()
            self.identify_subwell(SeBaView_wrapper,str(date_set_var),str(crystal_screen_var),str(target_protein_var),str(target_protein_top_left_stock_concentration_var),str(target_protein_top_right_stock_concentration_var),str(target_protein_bottom_left_stock_concentration_var),str(crystallization_chaperone_var),str(custom_tags_values))

    def New_Tray(self):
        self.clear_widgets()
        self.add_menu()
        self.root.geometry()
        new_tray_frame = ttk.Frame(self.root,padding="3 3 12 12")
        new_tray_frame.grid(column=0,row=0,sticky=(N,W))
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)

        ttk.Label(new_tray_frame, text="Select from standard tags or type a new entry:").grid(column=1,row=1)

        date_set_values = ["01.01.2025"] #Replace with code that accesses a page in an excel workbook that contains the date_set_values of each tray in the CrystalDex.
        date_set_label = ttk.Label(new_tray_frame,text="Date Set (required; 00.00.0000):")
        date_set_label.grid(column=1,row=5,sticky=(N,W))
        date_set_var = StringVar()
        date_set_drop_down = ttk.Combobox(new_tray_frame,textvariable=date_set_var,values=date_set_values)
        date_set_drop_down.grid(column=2,row=5)

        crystallization_chaperone_label = ttk.Label(new_tray_frame,text="Crystallization Chaperone (optional):")
        crystallization_chaperone_label.grid(column=1,row=6,sticky=(N,W))
        crystallization_chaperone_var = StringVar()
        crystallization_chaperone_drop_down = ttk.Combobox(new_tray_frame,textvariable=crystallization_chaperone_var,values=self.crystallization_chaperone_values)
        crystallization_chaperone_drop_down.grid(column=2,row=6)

        crystal_screen_label = ttk.Label(new_tray_frame,text="Crystal Screen (required):")
        crystal_screen_label.grid(column=1,row=7,sticky=(N,W))
        crystal_screen_var = StringVar()
        crystal_screen_drop_down = ttk.Combobox(new_tray_frame,textvariable=crystal_screen_var,values=self.crystal_screen_values)
        crystal_screen_drop_down.grid(column=2,row=7)

        target_protein_values = ["DARPin","CMG2","UBA","TELSAM","sfGFP"]
        target_protein_label = ttk.Label(new_tray_frame,text="Target Protein (required):")
        target_protein_label.grid(column=1,row=8,sticky=(N,W))
        target_protein_var = StringVar()
        target_protein_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_var,values=target_protein_values)
        target_protein_drop_down.grid(column=2,row=8,sticky=(N,W))

        target_protein_stock_concentration_values = [1,5,15,20]
        target_protein_top_left_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into top left subwell (required):")
        target_protein_top_left_stock_concentration_label.grid(column=1,row=9,sticky=(N,W))
        target_protein_top_left_stock_concentration_var = DoubleVar()
        target_protein_top_left_stock_concentration_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_top_left_stock_concentration_var,values=target_protein_stock_concentration_values)
        target_protein_top_left_stock_concentration_drop_down.grid(column=2,row=9,sticky=(N,W))

        target_protein_top_right_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into top right subwell (required):")
        target_protein_top_right_stock_concentration_label.grid(column=1,row=10,sticky=(N,W))
        target_protein_top_right_stock_concentration_var = DoubleVar()
        target_protein_top_right_stock_concentration_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_top_right_stock_concentration_var,values=target_protein_stock_concentration_values)
        target_protein_top_right_stock_concentration_drop_down.grid(column=2,row=10,sticky=(N,W))

        target_protein_bottom_left_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into bottom left subwell (required):")
        target_protein_bottom_left_stock_concentration_label.grid(column=1,row=11,sticky=(N,W))
        target_protein_bottom_left_stock_concentration_var = DoubleVar()
        target_protein_bottom_left_stock_concentration_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_bottom_left_stock_concentration_var,values=target_protein_stock_concentration_values)
        target_protein_bottom_left_stock_concentration_drop_down.grid(column=2,row=11,sticky=(N,W))

        #Later, if I have time, I'll want to add a little virtual replica in column 3 of a single well (with the four subwells) so that the user can see exactly what they're filling out, and each subwell will have the concentration appear as they fill it in.

        custom_tags_values = []
        custom_tags_label = ttk.Label(new_tray_frame,text="Custom Tags (optional; separated by commas, please!):")
        custom_tags_label.grid(column=1,row=12,sticky=(N,W))
        custom_tags_var = StringVar()
        custom_tags_drop_down = ttk.Combobox(new_tray_frame,textvariable=custom_tags_var,values=custom_tags_values)
        custom_tags_drop_down.grid(column=2,row=12)

        ttk.Button(new_tray_frame,text="Begin Indexing Tray",
                   command=lambda: self.Index_Tray(
                       date_set_var.get(),
                       crystal_screen_var.get(),
                       target_protein_var.get(),
                       target_protein_top_left_stock_concentration_var.get(),
                       target_protein_top_right_stock_concentration_var.get(),
                       target_protein_bottom_left_stock_concentration_var.get(),
                       crystallization_chaperone_var.get(),
                       custom_tags_var.get()
                        )
                    ).grid(column=1,row=13,sticky=(N,W))

        for child in new_tray_frame.winfo_children():
            child.grid_configure(padx=5,pady=5)

    def Open_Tray(self):
        self.clear_widgets()
        self.add_menu()
        open_tray_frame = ttk.Frame(self.root, padding="3 3 12 12").grid(column=0,row=0,sticky=(N,W,E,S))

    def Upload_Xtal_Screen(self):
        self.clear_widgets()
        self.add_menu()
        upload_xtal_screen_frame = ttk.Frame(self.root,padding="3 3 12 12").grid(column=0,row=0,sticky=(N,W,E,S))

    def startup(self):
        self.clear_widgets()
        self.add_menu()
        startup = ttk.Frame(self.root,padding='5 5 20 20')
        startup.option_add('*tearOFF',FALSE)
        startup.grid(column=0,row=0,sticky='N,E,S,W')
        #To make the buttons bigger and prettier, you'll have to use another widget, probably a text widget with a button placed inside it.
        #https://tkdocs.com/tutorial/text.html#basics
        new_tray_button = ttk.Button(startup,text="Index New Tray",command=self.New_Tray,width=40).grid(column=0,row=0,padx=50,pady=50,sticky=(N,E,S,W))
        open_tray_button = ttk.Button(startup,text="Open Tray",command=self.Open_Tray,width=40).grid(column=1,row=0,padx=50,pady=50,sticky=(N,E,S,W))
        upload_crystallization_screen_button = ttk.Button(startup,text="Upload Crystallization Screen",command=self.Upload_Xtal_Screen,width=40).grid(column=2,row=0,padx=50,pady=50,sticky=(N,E,S,W))
        self.root.mainloop()

if __name__ == "__main__":
    app = CrystalDex_main()
    app.startup()
