#See the README for the functional goals of this program and for author notes and acknowledgements.
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
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        self.root.minsize(self.screen_width//5,600)
        self.root.geometry(f'1050x700+{self.screen_width//2-525}+{self.screen_height//2-350}')
        #Make the window resizable:
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)
        self.crystallization_chaperone_values = ["1TEL","2TEL","3TEL","4TEL","5TEL","6TEL"]
        self.crystal_screen_values = ["Crystal_Screen","Index","PEG_Custom","PEG_Ion","Salt_Rx","Wizard"]
        token = 'rSiKvENhVt1CGdzsPp7oo8lUskXBLhUK'
        auth: BoxDeveloperTokenAuth = BoxDeveloperTokenAuth(token=token)
        self.client: BoxClient = BoxClient(auth=auth)

    def Box_Save(self):
        #The following command uploads the Crystal Trays Library to Box.
        self.client.uploads.upload_file_version(
            attributes=box_sdk_gen.UploadFileAttributesParentField(
                name="Crystal_Trays_Library.xlsx",
                id="1862599427539"),
                file_id="1862599427539",
                file=open(os.path.abspath("Crystal_Trays_Library.xlsx"),"rb"
                )
            )
    
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
            """
                #First, kill any current SeBaView window, then proceed.
                for proc in psutil.process_iter(attrs=["pid", "name", "exe"]):
                    try:
                        if proc.info["name"] and "SeBaView" in proc.info["name"]:
                            print(f"Killing SeBaView process: PID {proc.info['pid']}")
                            proc.kill()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        continue
                time.sleep(2)
            """
            with open("SeBaView_path_file.json", "w") as s:
                json.dump({"SeBaView_path": exe_path}, s)
            SeBaView = Application(backend="uia").start(exe_path)
            time.sleep(5)
            for i, w in enumerate(SeBaView.windows()):
                print(f'[{i}] Title: {w.window_text()}')
            try:
                # Try to wait for the first active window
                SeBaView_main_window = SeBaView.window(title_re=".*SeBaView.*")
                #timings.wait_until_passes(15, 1, lambda: main_window.exists() and main_window.is_visible())
                SeBaView_wrapper = SeBaView_main_window.wrapper_object()
                SeBaView_wrapper.set_focus()
                SeBaView_wrapper.maximize()
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

    def identify_subwell(self,ws,SeBaView_wrapper,date_set_var,crystal_screen_var,target_protein_var,target_protein_top_left_stock_concentration_var,target_protein_top_right_stock_concentration_var,target_protein_bottom_left_stock_concentration_var,crystallization_chaperone_var,custom_tags_values):
        self.clear_widgets()
        self.add_menu()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry(f"{screen_width // 4}x{screen_height}+0+0")
        subwell_frame = ttk.Frame(self.root,padding="3 3 12 12")
        subwell_frame.grid(column=0,row=0,sticky=(N,W))
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)

        ensure_magnified_label = ttk.Label(subwell_frame,text="MAKE SURE the microscope is fully\nmagnified before taking any pictures.")
        ensure_magnified_label.grid(column=1,row=1)

        well_column_label = ttk.Label(subwell_frame,text="Well column:")
        well_column_label.grid(column=1,row=2)
        well_column_values = ['A','B','C','D','E','F','G','H']
        well_column_var = StringVar()
        well_column_drop_down = ttk.Combobox(subwell_frame,textvariable=well_column_var,values=well_column_values)
        well_column_drop_down.grid(column=2,row=2)

        well_row_label = ttk.Label(subwell_frame,text="Well row:")
        well_row_label.grid(column=1,row=3)
        well_row_values = [1,2,3,4,5,6,7,8,9,10,11,12]
        well_row_var = IntVar()
        well_row_drop_down = ttk.Combobox(subwell_frame,textvariable=well_row_var,values=well_row_values)
        well_row_drop_down.grid(column=2,row=3)

        subwell_values = ['top_left','top_right','bottom_left']
        subwell_label = ttk.Label(subwell_frame,text="subwell:")
        subwell_label.grid(column=1,row=4,sticky=(N,W))
        subwell_var = StringVar()
        subwell_drop_down = ttk.Combobox(subwell_frame,textvariable=subwell_var,values=subwell_values)
        subwell_drop_down.grid(column=2,row=4)

        number_of_crystals_label = None
        number_of_crystals_var = None

        possible_salt_crystals_label = ttk.Label(subwell_frame,text="Possibly a salt crystal")
        possible_salt_crystals_label.grid(column=1,row=5)
        possible_salt_crystals_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=possible_salt_crystals_var,onvalue=True,offvalue=False).grid(column=2,row=5)

        precipitation_label = ttk.Label(subwell_frame,text="Precipitation present")
        precipitation_label.grid(column=1,row=6)
        precipitation_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=precipitation_var,onvalue=True,offvalue=False).grid(column=2,row=6)

        microcrystals_label = ttk.Label(subwell_frame,text="Microcrystals present")
        microcrystals_label.grid(column=1,row=7)
        microcrystals_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=microcrystals_var,onvalue=True,offvalue=False).grid(column=2,row=7)

        glassy_protein_or_artifacts_label = ttk.Label(subwell_frame,text="Glassy protein or artifacts present")
        glassy_protein_or_artifacts_label.grid(column=1,row=8)
        glassy_protein_or_artifacts_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=glassy_protein_or_artifacts_var,onvalue=True,offvalue=False).grid(column=2,row=8)

        now = datetime.now()
        date_snapped = now.strftime('%m-%d-%Y')
        image_title = f'{crystallization_chaperone_var}_{target_protein_var}_{crystal_screen_var}_{well_column_var.get()}{well_row_var.get()}_{subwell_var.get()}_{date_set_var}_{date_snapped}' 

        ttk.Button(subwell_frame,text="Take and Save Picture",
                command=lambda: self.take_picture(
                     SeBaView_wrapper,
                     image_title,
                     ws
                )).grid(column=1,row=14,sticky=(N,W))
        
        for child in subwell_frame.winfo_children():
            child.grid_configure(padx=15,pady=15)

        #Refocus the window if minimized
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

        #This needs to be updated, and the Mastercopy needs to be edited as well to fill in all the necessary information.
        well_to_excel_dict = {'A':8,'B':19,'C':30,'D':41,'E':52,'F':63,'G':74,'H':85,1:'C',2:'H',3:'M',4:'R',5:'W',6:'AB',7:'AG',8:'AL',9:'AQ',10:'AV',11:'BA',12:'BF'}
        picture_link_cell = ws[f'{well_to_excel_dict[f'{well_column_var.get()}']}{well_to_excel_dict[f'{well_row_var.get()}']}']
        picture_link_cell.value(f'{image_title}')
        #picture_link_cell.hyperlink()#Insert the hyperlink value into this.
        #picture_link_cell.offset(row=0,column=1).value(f'{number_of_crystals_var.get()}')#This is how to fill other cells as a reference of the picture_link_cell

    def take_picture(self,SeBaView_wrapper,image_title,ws):
        SeBaView_wrapper.set_focus() #Is this needed? Test later...
        SeBaView_wrapper.click_input(coords=(55, 70))  #This accesses the save as button.
        time.sleep(3)
        pywinauto.keyboard.send_keys(f"{image_title}{{ENTER}}") #I believe this code will let me type into whatever location is currently focused, but I haven't tested it yet.
        for filename in os.listdir(downloads):
            file_path = os.path.join(downloads, filename)
            if os.path.isfile(file_path) and filename.lower().endswith(('.jpeg','.jpg')) and os.path.getmtime(file_path)<1:
                try:
                    shutil.move(file_path, crystal_pictures)
                    print(f"Moved: {filename}")
                except Exception as e:
                    print(f"Failed to move {filename}: {e}")
        self.Box_Save()

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

    def proceed(self,ws):
        SeBaView_wrapper = self.load_SeBaView()
        date_set_var = ws['D1']
        crystal_screen_var = ws['D3']
        target_protein_var = ws['D4']
        target_protein_top_left_stock_concentration_var = ws['H1']
        target_protein_top_right_stock_concentration_var = ws['H2']
        target_protein_bottom_left_stock_concentration_var = ws['H3']
        crystallization_chaperone_var = ws['D2']
        custom_tags_values = ws['D5']
        self.identify_subwell(ws,SeBaView_wrapper,date_set_var,crystal_screen_var,target_protein_var,target_protein_top_left_stock_concentration_var,target_protein_top_right_stock_concentration_var,target_protein_bottom_left_stock_concentration_var,crystallization_chaperone_var,custom_tags_values)

    def Select_Tray(self,wb,tray_names):
        self.clear_widgets()
        self.add_menu()
        self.root.geometry()
        select_tray_frame = ttk.Frame(self.root,padding="3 3 12 12")
        select_tray_frame.grid(column=0,row=0,sticky=(N,W))
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)

        select_tray_name_label = ttk.Label(select_tray_frame,text=(
            'At least one previously indexed tray was found that shares a date, screen, and target protein with the current tray.'
            '\nPlease review the following to ensure no duplicate trays are indexed!'
        ))
        select_tray_name_label.grid(column=0,row=0)
        tray_name = StringVar()
        select_tray_name_combobox = ttk.Combobox(select_tray_frame,values=tray_names,textvariable=tray_name)
        select_tray_name_combobox.grid(column=0,row=1)

        none_of_the_above_var = BooleanVar()
        none_of_the_above_label = ttk.Label(select_tray_frame,text='If none of the above match your tray, click here:')
        none_of_the_above_label.grid(column=0,row=2)
        none_of_the_above_checkbutton = ttk.Checkbutton(select_tray_frame,variable=none_of_the_above_var,onvalue=True,offvalue=False)
        none_of_the_above_checkbutton.grid(column=1,row=2)
        if none_of_the_above_var:
            tray_name.set("")

        ttk.Button(select_tray_frame,text="Save selection and proceed",
        command=lambda: self.proceed(wb[str(tray_name.get())])).grid(column=1,row=14,sticky=(N,W))

    def Index_Tray(self,date_set_var,crystal_screen_var,target_protein_var,target_protein_top_left_stock_concentration_var,target_protein_top_right_stock_concentration_var,target_protein_bottom_left_stock_concentration_var,crystallization_chaperone_var,custom_tags_values):
        indexable = False
        file_download = None
        #Open up the current Box version of the Crystal_Trays_Library.xlsx or the mastercopy if none exists:
        try:
            file_id = '1862599427539'
            file_download = self.client.downloads.download_file(file_id).read()
        except FileNotFoundError:
            file_id = '1861370891462'
            file_download = self.client.downloads.download_file(file_id).read()

        #Begin writing a new Crystal_Trays_Library.xlsx file on the working computer:
        with open("Crystal_Trays_Library.xlsx","wb") as c:
            c.write(file_download)
        print("Saving to:", os.path.abspath("Crystal_Trays_Library.xlsx"))
        wb = load_workbook(filename=os.path.abspath("Crystal_Trays_Library.xlsx")) #Accesses the excel workbook using the openpyxl python module
        date = None
        crystal_screen = None
        target_protein = None
        try:
            date = str(datetime.strftime(datetime.strptime(date_set_var,"%m-%d-%Y"),"%m-%d-%Y"))
            print(f'indexing! date: {date_set_var}')
            indexable = True
        except ValueError:
            messagebox.showerror(title="Date Error",message="You attempted to put in an invalid date. Please use the style: 01-01-2025")
            print(f'not indexing. date: {date_set_var}')
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
        custom_tags_list = [tag.strip() for tag in custom_tags_values.split(', ')]
        if indexable:
            ws_possible_duplicate_count = 0
            for ws in wb:
                tags_cell = str(ws['K1'].value or "")
                tags = [tag.strip() for tag in tags_cell.split(', ')]
                print(f'tags: {tags}, [date,crystal_screen,target_protein]: {[date,crystal_screen,target_protein]}')
                if all(term in tags for term in [date,crystal_screen,target_protein]):
                    ws_possible_duplicate_count += 1
            if ws_possible_duplicate_count >0:
                tray_names = []
                for ws in wb:
                    if all(term in ws.title for term in [date,crystal_screen,target_protein]):
                        tray_names.append(ws.title)
                print(f'tray_names: {tray_names}')
                self.Select_Tray(wb,tray_names)
            elif ws_possible_duplicate_count == 0:
                print(f"No trays found with these stats; generating new tray!")
                new_worksheet = wb.copy_worksheet(wb["Mastercopy"])
                full_title = f'{target_protein}_{crystal_screen}_{date}_1'
                short_title = full_title[:26]
                print(f'short_title: {short_title}')
                new_worksheet.title = short_title
                all_tags = [date_set_var,crystal_screen_var,target_protein_var,target_protein_top_left_stock_concentration_var,target_protein_top_right_stock_concentration_var,target_protein_bottom_left_stock_concentration_var,crystallization_chaperone_var,custom_tags_values]
                new_worksheet['K1'] = ', '.join(map(str,all_tags))
                new_worksheet['D1'] = date_set_var
                new_worksheet['D2'] = crystallization_chaperone_var
                new_worksheet['D3'] = crystal_screen_var
                new_worksheet['D4'] = target_protein_var
                new_worksheet['D5'] = custom_tags_values
                new_worksheet['H1'] = target_protein_top_left_stock_concentration_var
                new_worksheet['H2'] = target_protein_top_right_stock_concentration_var
                new_worksheet['H3'] = target_protein_bottom_left_stock_concentration_var
                wb.save(filename=os.path.abspath("Crystal_Trays_Library.xlsx"))
                SeBaView_wrapper = self.load_SeBaView()
                self.identify_subwell(new_worksheet,SeBaView_wrapper,date_set_var,crystal_screen_var,target_protein_var,target_protein_top_left_stock_concentration_var,target_protein_top_right_stock_concentration_var,target_protein_bottom_left_stock_concentration_var,crystallization_chaperone_var,custom_tags_values)

    def New_Tray(self):
        self.clear_widgets()
        self.add_menu()
        self.root.geometry()
        new_tray_frame = ttk.Frame(self.root,padding="3 3 12 12")
        new_tray_frame.grid(column=0,row=0,sticky=(N,W))
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)

        ttk.Label(new_tray_frame, text="Select from standard tags or type a new entry:").grid(column=1,row=1)

        date_set_values = [str(datetime.now().strftime('%m-%d-%Y'))] #Replace with code that accesses a page in an excel workbook that contains the date_set_values of each tray in the CrystalDex.
        today_label = ttk.Label(new_tray_frame,text="Today?")
        today_label.grid(column=3,row=5,sticky=(N,W))
        today_var = BooleanVar()
        ttk.Checkbutton(new_tray_frame,variable=today_var,onvalue=True,offvalue=False).grid(column=4,row=5,sticky=(N,W))
        date_set_label = ttk.Label(new_tray_frame,text="Date Set (required; 00-00-0000):")
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
        if today_var:
            date_set_var.set(datetime.now().strftime('%m-%d-%Y'))
            print(f'date_set_var.get() = {date_set_var.get()}')

        ttk.Button(new_tray_frame,text="Begin Indexing Tray",
                   command=lambda: self.Index_Tray(
                       str(date_set_var.get()),
                       str(crystal_screen_var.get()),
                       str(target_protein_var.get()),
                       str(target_protein_top_left_stock_concentration_var.get()),
                       str(target_protein_top_right_stock_concentration_var.get()),
                       str(target_protein_bottom_left_stock_concentration_var.get()),
                       str(crystallization_chaperone_var.get()),
                       str(custom_tags_var.get())
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
        self.root.geometry(f'1050x700+{self.screen_width//2-525}+{self.screen_height//2-350}')
        startup.option_add('*tearOFF',FALSE)
        startup.grid(column=0,row=0,sticky='N,E,S,W')
        #To make the buttons bigger and prettier, you'll have to use another widget, probably a text widget with a button placed inside it.
        #https://tkdocs.com/tutorial/text.html#basics
        ttk.Button(startup,text="Index New Tray",command=self.New_Tray,width=40).grid(column=0,row=0,padx=50,pady=50,sticky=(N,E,S,W))
        ttk.Button(startup,text="Open Tray",command=self.Open_Tray,width=40).grid(column=1,row=0,padx=50,pady=50,sticky=(N,E,S,W))
        ttk.Button(startup,text="Upload Crystallization Screen",command=self.Upload_Xtal_Screen,width=40).grid(column=2,row=0,padx=50,pady=50,sticky=(N,E,S,W))
        self.root.mainloop()

if __name__ == "__main__":
    app = CrystalDex_main()
    app.startup()
