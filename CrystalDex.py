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
from box_sdk_gen import BoxClient, BoxOAuth, OAuthConfig, FileTokenStorage, BoxSDKError, UploadFileAttributes, UploadFileAttributesParentField
import webbrowser

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
script_dir = os.path.dirname(os.path.abspath(__file__))#The directory of this script, so basically the folder where all the code is kept.
icon_path = os.path.join(script_dir,"crystaldex_icon.png")
crystal_pictures = os.path.join(script_dir,"Crystal_Pictures")
os.makedirs(crystal_pictures,exist_ok=True)

home = os.path.expanduser("~")
downloads = os.path.join(home, "Downloads")
desktop = os.path.expanduser("~/Desktop")

def on_click(x,y,button,pressed):
    global mouse_is_down
    mouse_is_down = pressed

listener = mouse.Listener(on_click=on_click)
listener.start()

class CrystalDex_main:
    def __init__(self):
        #Box integrations:
        CLIENT_ID = "ywdxl21bfyxj6lpzest9alondci3jezf"
        CLIENT_SECRET = "WV4AhaJ4P0b6UHy8ENaXTNby6mjyxJv5"
        token_storage = FileTokenStorage(filename='box_token.json')
        config = OAuthConfig(
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET,
            token_storage=token_storage
        )
        auth = BoxOAuth(config)
        try:
            auth.retrieve_token()
            print(f'User already approved app for Box.')
        except BoxSDKError:
            auth_url = auth.get_authorize_url()
            webbrowser.open(auth_url)
            authorization_code = input("Paste the code you got after approving: ")
            auth.get_tokens_authorization_code_grant(authorization_code)
        self.client = BoxClient(auth=auth)

        #Access CrystalDex_Library or the Mastercopy:
        file_download = None
        try:
            file_id = '1892938696722'
            file_download = self.client.downloads.download_file(file_id).read()
        except FileNotFoundError:
            file_id = '1861370891462'
            file_download = self.client.downloads.download_file(file_id).read()

        #Begin writing a new CrystalDex_Library.xlsx file on the working computer:
        with open("CrystalDex_Library.xlsx","wb") as c:
            c.write(file_download)
        print("Saving to:", os.path.abspath("CrystalDex_Library.xlsx"))
        self.wb = load_workbook(filename=os.path.abspath("CrystalDex_Library.xlsx")) #Accesses the excel workbook using the openpyxl python module
        
        #Other frequently accessed values (will be turned into a .json soon):
        self.crystallization_chaperone_values = ["1TEL","2TEL","3TEL","4TEL","5TEL","6TEL"]
        self.crystal_screen_values = ["Crystal_Screen","Index","PEG_Custom","PEG_Ion","Salt_Rx","Wizard"]
        self.crystal_screen_symbols = {'Crystal_Screen':'CS',
                                       'Index':'IN',
                                       'PEG_Custom':'PC',
                                       'PEG_Ion':'PI',
                                       'Salt_Rx':'SR',
                                       'Wizard':'WI'}
        self.crystal_size = [0,0]
        self.pixel_to_size = 1000/458 #1 millimeter or 1000 microns per 458 pixels at 40% magnification (ie. a picture size of 2560x1922pixels on the screen)
        self.picture_upload_paths = []
        self.button_location = None

        #Tkinter initializations
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

    def refocus(self):
        #Refocus the window if minimized
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def Box_Save(self):
        #The following command uploads the Crystal Trays Library and the images to Box.
        self.client.uploads.upload_file_version(
            attributes=box_sdk_gen.UploadFileAttributesParentField(
                name="CrystalDex_Library.xlsx",
                id="320928486478"),
                file_id="1892938696722",
                file=open(os.path.abspath("CrystalDex_Library.xlsx"),"rb"
                )
            )
        for image_filename in os.listdir(crystal_pictures):
            file_path = os.path.join(crystal_pictures, image_filename)
            with open(file_path,'rb') as image_stream:
                uploading_file_return = self.client.uploads.upload_file(
                    UploadFileAttributes(
                        name=image_filename,parent=UploadFileAttributesParentField(id='325857937585')#The id here is where the images will end up.
                    ),
                    image_stream
                )
                uploading_file = uploading_file_return.entries[0]
                self.client.shared_links_files.add_share_link_to_file(file_id=uploading_file.id,fields='shared_link')
                shared_link_id = self.client.shared_links_files.get_shared_link_for_file(uploading_file.id,"shared_link").id
                shared_link_url = 'https://byu.app.box.com/file/'+str(shared_link_id)
                print(f'shared_link: {shared_link_url}')
                #Add code here to find the correct excel sheet for the file and add the link to it...
        self.startup()
    
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
            with open("SeBaView_path_file.json", "w") as s:
                json.dump({"SeBaView_path": exe_path}, s)
            self.SeBaView = Application(backend="uia").start(exe_path)
            time.sleep(5)
            for i, w in enumerate(self.SeBaView.windows()):
                print(f'[{i}] Title: {w.window_text()}')
            try:
                # Try to wait for the first active window
                SeBaView_main_window = self.SeBaView.window(title_re=".*SeBaView.*")
                #timings.wait_until_passes(15, 1, lambda: main_window.exists() and main_window.is_visible())
                SeBaView_wrapper = SeBaView_main_window.wrapper_object()
                SeBaView_wrapper.set_focus()
                SeBaView_wrapper.maximize()
                
                print("SeBaView main window focused successfully.")
            except Exception as e:
                print("Failed to find or focus the SeBaView window. Restarting your computer usually fixes this issue.")
                print(f"Error: {e}")

            for i in range(2):
                SeBaView_wrapper.click_input(coords=(60, 165))  #This accesses the camera connecting button.
            return SeBaView_wrapper

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
        ttk.Button(startup,text="Edit Tray",command=self.Edit_Tray,width=40).grid(column=1,row=0,padx=50,pady=50,sticky=(N,E,S,W))
        ttk.Button(startup,text='Harvest Crystals',command=self.Harvest_Crystals,width=40).grid(column=2,row=0,padx=50,pady=50,sticky=(N,E,S,W))
        ttk.Button(startup,text="Upload Crystallization Screen",command=self.Upload_Xtal_Screen,width=40).grid(column=0,row=1,padx=50,pady=50,sticky=(N,E,S,W))
        self.root.mainloop()

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

        target_protein_stock_concentration_values = ['1','5','15','20']
        target_protein_top_left_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into top left subwell (required):")
        target_protein_top_left_stock_concentration_label.grid(column=1,row=9,sticky=(N,W))
        target_protein_top_left_stock_concentration_var = StringVar()
        target_protein_top_left_stock_concentration_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_top_left_stock_concentration_var,values=target_protein_stock_concentration_values)
        target_protein_top_left_stock_concentration_drop_down.grid(column=2,row=9,sticky=(N,W))

        target_protein_top_right_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into top right subwell (required):")
        target_protein_top_right_stock_concentration_label.grid(column=1,row=10,sticky=(N,W))
        target_protein_top_right_stock_concentration_var = StringVar()
        target_protein_top_right_stock_concentration_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_top_right_stock_concentration_var,values=target_protein_stock_concentration_values)
        target_protein_top_right_stock_concentration_drop_down.grid(column=2,row=10,sticky=(N,W))

        target_protein_bottom_left_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into bottom left subwell (required):")
        target_protein_bottom_left_stock_concentration_label.grid(column=1,row=11,sticky=(N,W))
        target_protein_bottom_left_stock_concentration_var = StringVar()
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
                       str(date_set_var.get()),
                       bool(today_var.get()),
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

    def Select_Tray(self,tray_names=None,short_title=None):
        self.clear_widgets()
        self.add_menu()
        self.root.geometry()
        select_tray_frame = ttk.Frame(self.root,padding="3 3 12 12")
        select_tray_frame.grid(column=0,row=0,sticky=(N,W))
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)
        if tray_names == None:
            tray_names = []
            for ws in self.wb:
                tray_names.append([ws.title,ws['A1']])

        select_tray_name_label = ttk.Label(select_tray_frame,text=(
            'At least one previously indexed tray was found that shares a date, screen, and target protein with the current tray.'
            '\nPlease review the following to ensure no duplicate trays are indexed!'
        ))
        select_tray_name_label.grid(column=0,row=0)
        tray_name = StringVar()
        select_tray_name_combobox = ttk.Combobox(select_tray_frame,values=tray_names,textvariable=tray_name)
        select_tray_name_combobox.grid(column=0,row=1)

        none_of_the_above_label = ttk.Label(select_tray_frame,text='If none of the above match your tray, click here:')
        none_of_the_above_label.grid(column=0,row=2)
        ttk.Button(select_tray_frame,text="make new tray",command=lambda: make_new_tray(short_title)).grid(column=1,row=2)

        def make_new_tray(short_title):
            new_worksheet = self.wb.copy_worksheet(self.wb["Mastercopy"])
            new_worksheet.title = short_title
            #This proceed thing is wrong. It needs to instead copy what is done after a normal tray is made...
            self.proceed(self.wb[short_title])

        ttk.Button(select_tray_frame,text="Save selection and proceed",
        command=lambda: self.proceed(self.wb[str(tray_name.get())])).grid(column=0,row=3)

    def Edit_Tray(self):
        self.clear_widgets()
        self.add_menu()
        edit_tray_frame = ttk.Frame(self.root, padding="3 3 12 12").grid(column=0,row=0,sticky=(N,W,E,S))

    def Index_Tray(self,date_set,today,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,crystallization_chaperone,custom_tags_values):
        print(f'values passed in: {date_set}, {crystal_screen}, {target_protein}, {target_protein_top_left_stock_concentration}, {target_protein_top_right_stock_concentration}, {target_protein_bottom_left_stock_concentration}, {crystallization_chaperone}, {custom_tags_values}')
        indexable = False
        if today:
            date_set = (datetime.now().strftime('%m-%d-%Y'))
        else:
            try:
                date = str(datetime.strftime(datetime.strptime(date_set,"%m-%d-%Y"),"%m-%d-%Y"))
                print(f'indexing! date: {date_set}')
                indexable = True
            except ValueError:
                messagebox.showerror(title="Date Error",message="You attempted to put in an invalid date. Please use the style: 01-01-2025")
                print(f'not indexing. date: {date_set}')
        if crystal_screen in self.crystal_screen_values:
            indexable = True
        else: 
            messagebox.showerror(title="Crystal Screen Does Not Exist",message="The crystal screen you attempted to reference does not exist.")
            indexable = False
        if target_protein is not None:
            indexable = True
        else:
            messagebox.showerror(title="No Protein Target",message="You neglected to enter a protein target. (CrystalDex can't index nothingness!)")
            indexable = False
        custom_tags_list = [tag.strip() for tag in custom_tags_values.split(', ')]
        if indexable:
            ws_possible_duplicate_count = 0
            for ws in self.wb:
                tags_cell = str(ws['K1'].value or "")
                tags = [tag.strip() for tag in tags_cell.split(', ')]
                print(f'tags: {tags}, [date,crystal_screen,target_protein]: {[date,crystal_screen,target_protein]}')
                if all(term in tags for term in [date,crystal_screen,target_protein]):
                    ws_possible_duplicate_count += 1
            if ws_possible_duplicate_count >0:
                tray_names = []
                for ws in self.wb:
                    tags_cell = str(ws['K1'].value or "")
                    tags = [tag.strip() for tag in tags_cell.split(', ')]
                    if all(term in tags for term in [date,crystal_screen,target_protein]):
                        tray_names.append(ws.title)
                print(f'tray_names: {tray_names}')
                full_title = f'{date}_{self.crystal_screen_symbols.get(crystal_screen)}_{target_protein}_1'
                short_title = full_title[:26]
                self.Select_Tray(tray_names,short_title)
            elif ws_possible_duplicate_count == 0:
                print(f"No trays found with these stats; generating new tray!")
                new_worksheet = self.wb.copy_worksheet(self.wb["Mastercopy"])
                full_title = f'{date}_{self.crystal_screen_symbols.get(crystal_screen)}_{target_protein}_1'
                short_title = full_title[:26]
                print(f'short_title: {short_title}')
                new_worksheet.title = short_title
                all_tags = [date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,crystallization_chaperone,custom_tags_values]
                new_worksheet['A1'] = full_title
                new_worksheet['K1'] = ', '.join(map(str,all_tags))
                new_worksheet['D1'] = date_set
                new_worksheet['D2'] = crystallization_chaperone
                new_worksheet['D3'] = crystal_screen
                new_worksheet['D4'] = target_protein
                new_worksheet['D5'] = custom_tags_values
                new_worksheet['H1'] = target_protein_top_left_stock_concentration
                new_worksheet['H2'] = target_protein_top_right_stock_concentration
                new_worksheet['H3'] = target_protein_bottom_left_stock_concentration
                self.wb.save(filename=os.path.abspath("CrystalDex_Library.xlsx"))
                SeBaView_wrapper = self.load_SeBaView()
                self.identify_subwell(new_worksheet,SeBaView_wrapper,date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,crystallization_chaperone,custom_tags_values)

    def identify_subwell(self,ws,SeBaView_wrapper,date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,crystallization_chaperone,custom_tags_values):
        self.clear_widgets()
        self.add_menu()
        self.root.geometry(f"{self.screen_width // 4}x{self.screen_height}+0+0")
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
        well_row_values = ['1','2','3','4','5','6','7','8','9','10','11','12']
        well_row_var = StringVar()
        well_row_drop_down = ttk.Combobox(subwell_frame,textvariable=well_row_var,values=well_row_values)
        well_row_drop_down.grid(column=2,row=3)

        subwell_values = ['top_left','top_right','bottom_left']
        subwell_label = ttk.Label(subwell_frame,text="subwell:")
        subwell_label.grid(column=1,row=4)
        subwell_var = StringVar()
        subwell_drop_down = ttk.Combobox(subwell_frame,textvariable=subwell_var,values=subwell_values)
        subwell_drop_down.grid(column=2,row=4)

        crystal_width_var = StringVar()
        crystal_width_label = ttk.Label(subwell_frame,text='crystal width:')
        crystal_width_label.grid(column=1,row=5)
        crystal_width_entry = ttk.Entry(subwell_frame,textvariable=crystal_width_var,state=DISABLED)
        crystal_width_entry.grid(column=2,row=5)
        um_width_label = ttk.Label(subwell_frame,text='um')
        um_width_label.grid(column=3,row=5)

        crystal_height_var = StringVar()
        crystal_height_label = ttk.Label(subwell_frame,text='crystal height:')
        crystal_height_label.grid(column=1,row=6)
        crystal_height_entry = ttk.Entry(subwell_frame,textvariable=crystal_height_var,state=DISABLED)
        crystal_height_entry.grid(column=2,row=6)
        um_row_label = ttk.Label(subwell_frame,text='um')
        um_row_label.grid(column=3,row=6)

        number_of_crystals_label = ttk.Label(subwell_frame,text='# of harvestable crystals (optional):')
        number_of_crystals_label.grid(column=1,row=7)
        number_of_crystals_var = StringVar()
        number_of_crystals_entry = ttk.Spinbox(subwell_frame,from_=0,to=100,textvariable=number_of_crystals_var)
        number_of_crystals_entry.grid(column=2,row=7)

        shape_label = ttk.Label(subwell_frame,text='Shape of crystals (optional):')
        shape_label.grid(column=1,row=8)
        shape_var = StringVar()
        shape_entry = ttk.Entry(subwell_frame,textvariable=shape_var)
        shape_entry.grid(column=2,row=8)

        possible_salt_crystals_label = ttk.Label(subwell_frame,text="Possibly a salt crystal")
        possible_salt_crystals_label.grid(column=1,row=9)
        possible_salt_crystals_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=possible_salt_crystals_var,onvalue=True,offvalue=False).grid(column=2,row=9)

        precipitation_label = ttk.Label(subwell_frame,text="Precipitation present")
        precipitation_label.grid(column=1,row=10)
        precipitation_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=precipitation_var,onvalue=True,offvalue=False).grid(column=2,row=10)

        microcrystals_label = ttk.Label(subwell_frame,text="Microcrystals present")
        microcrystals_label.grid(column=1,row=11)
        microcrystals_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=microcrystals_var,onvalue=True,offvalue=False).grid(column=2,row=11)

        glassy_protein_or_artifacts_label = ttk.Label(subwell_frame,text="Glassy protein or artifacts present")
        glassy_protein_or_artifacts_label.grid(column=1,row=12)
        glassy_protein_or_artifacts_var = BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=glassy_protein_or_artifacts_var,onvalue=True,offvalue=False).grid(column=2,row=12)

        now = datetime.now()
        date_snapped = now.strftime('%m-%d-%Y')

        notes_label = ttk.Label(subwell_frame,text="Crystallographer notes:")
        notes_label.grid(column=1,row=15)
        notes = Text(subwell_frame, width = 30, height = 10)
        notes.grid(column=1,row=16)

        def update_crystal_size_vars():
            crystal_width_var.set(f'{self.crystal_size[0]}')
            crystal_height_var.set(f'{self.crystal_size [1]}')

        ttk.Button(subwell_frame,text ='Measure Crystal',
                   command=lambda: self.measure_crystal(
                       SeBaView_wrapper,
                       update_crystal_size_vars
                   )).grid(column=1,row=13)

        #This function is having image_title issues...
        ttk.Button(subwell_frame,text="Take and Save Picture",
                command=lambda: self.take_picture(
                     SeBaView_wrapper,
                     ws,
                     crystallization_chaperone,
                     target_protein,
                     crystal_screen,
                     str(well_column_var.get()),
                     str(well_row_var.get()),
                     str(subwell_var.get()),
                     str(number_of_crystals_var.get()),
                     str(shape_var.get()),
                     bool(possible_salt_crystals_var.get()),
                     bool(precipitation_var.get()),
                     bool(microcrystals_var.get()),
                     bool(glassy_protein_or_artifacts_var.get()),
                     date_set,
                     date_snapped,
                     notes.get(1.0,END)
                )).grid(column=1,row=14)

        ttk.Button(subwell_frame,text="Done",
                   command=lambda: self.Box_Save()).grid(column=2,row=16)
        
        for child in subwell_frame.winfo_children():
            child.grid_configure(padx=5,pady=15)
        self.refocus()

    def proceed(self,ws):
        #This may need to be able to access subwells eventually...
        SeBaView_wrapper = self.load_SeBaView()
        date_set = ws['D1'].value
        crystal_screen = ws['D3'].value
        target_protein = ws['D4'].value
        target_protein_top_left_stock_concentration = ws['H1'].value
        target_protein_top_right_stock_concentration = ws['H2'].value
        target_protein_bottom_left_stock_concentration = ws['H3'].value
        crystallization_chaperone = ws['D2'].value
        custom_tags_values = ws['D5'].value
        self.identify_subwell(ws,SeBaView_wrapper,date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,crystallization_chaperone,custom_tags_values)

    def monitor_mouse(self,SeBaView_wrapper_rect):
        print(f'monitor_mouse is running...')
        if mouse_is_down:
            x, y = pyautogui.position()
            rel_x = x - SeBaView_wrapper_rect.left
            rel_y = y - SeBaView_wrapper_rect.top
            print(f'Mouse pressed relative to SeBaView window: ({rel_x}, {rel_y})')
            self.button_location = rel_x,rel_y
        time.sleep(0.1)
        print(f'monitor mouse is done.')

    def take_picture(
            self,
            SeBaView_wrapper,
            ws,
            crystallization_chaperone,
            target_protein,
            crystal_screen,
            well_column,
            well_row,
            subwell,
            number_of_crystals,
            shape,
            possible_salt_crystals,
            precipitation,
            microcrystals,
            glassy_protein_or_artifacts,
            date_set,
            date_snapped,
            notes):
        image_title = f'{crystallization_chaperone}_{target_protein}_{crystal_screen}_{well_column}{well_row}_{subwell}_{date_set}_{date_snapped}' 
        print(f'image_title: {image_title}')
        SeBaView_wrapper.set_focus() #Is this needed? Test later...
        SeBaView_wrapper.click_input(coords=(55, 70))  #This accesses the save as button.
        time.sleep(3)
        #The following three lines of code are to enable me to see where I need to program a mouse click to get the images to always save to the right spot.
        #SeBaView_wrapper_rect = SeBaView_wrapper.rectangle()
        #while self.button_location is None:
        #    self.monitor_mouse(SeBaView_wrapper_rect)
        for i in range(2):
            SeBaView_wrapper.click_input(coords=(750,450))#This is supposed to access the Desktop button to save the photos there temporarily, although it might not work. I may need to get a wrapper for the save window that opens... 
        pywinauto.keyboard.send_keys(f"{image_title}{{ENTER}}") #Enter the image_title name into the save window
        time.sleep(3)#sleep to let the file settle so the next command will be able to grab and move it

        for filename in os.listdir(desktop):
            print(f'filename: {filename}')
            file_path = os.path.join(desktop, filename)
            if os.path.isfile(file_path) and filename.lower().endswith(('.jpeg','.jpg','.bmp')) :#and time.time() - os.path.getmtime(file_path)<10
                try:
                    shutil.move(file_path, crystal_pictures)
                    self.picture_upload_filenames.append([filename])
                    print(f"Moved: {filename}")
                except Exception as e:
                    print(f"Failed to move {filename}: {e}")

        #This needs to be updated, and the Mastercopy needs to be edited as well to fill in all the necessary information.
        well_to_excel_dict = {'A':8,'B':23,'C':38,'D':53,'E':68,'F':83,'G':98,'H':113,'1':'C','2':'H','3':'M','4':'R','5':'W','6':'AB','7':'AG','8':'AL','9':'AQ','10':'AV','11':'BA','12':'BF'}
        row = well_to_excel_dict.get(well_row)
        column = well_to_excel_dict.get(well_column)
        
        picture_link_cell = ws[f'{row}{column}']
        if subwell=='top_right':
            picture_link_cell = picture_link_cell.offset(row=0,column=2)
        elif subwell=='bottom_left':
            picture_link_cell = picture_link_cell.offset(row=7,column=0)
        print(f'row: {picture_link_cell.row}, column: {picture_link_cell.column}')
        picture_link_cell.value = image_title
        picture_link_cell.offset(row=1,column=0).value = f'{(datetime.strptime(date_snapped,"%m-%d-%Y")-datetime.strptime(date_set,"%m-%d-%Y")).days}' #might have to change the type of these variables
        picture_link_cell.offset(row=2,column=0).value = f'{self.crystal_size[0]}x{self.crystal_size[1]} um'
        picture_link_cell.offset(row=3,column=0).value = f'{number_of_crystals}'
        picture_link_cell.offset(row=4,column=0).value = f'{shape}'
        if possible_salt_crystals:
            precipitation,
            microcrystals,
            glassy_protein_or_artifacts,
        picture_link_cell.offset(row=5,column=0).value = f'Possible salt crystals: {possible_salt_crystals}, precipitation: {precipitation}, microcrystals: {microcrystals}, glassy protein or artifacts: {glassy_protein_or_artifacts}'
        picture_link_cell.offset(row=6,column=0).value = notes

        #Move this code to the Box Save function because uploading and getting links is time-intensive. May be a puzzle!
        #picture_link_cell.hyperlink()#Insert the hyperlink value into this.
        #picture_link_cell.offset(row=0,column=1).value(f'{number_of_crystals_var.get()}')
        #self.wb.save(filename=os.path.abspath("CrystalDex_Library.xlsx"))
        
        self.wb.save(filename=os.path.abspath("CrystalDex_Library.xlsx"))
        if hasattr(self,'measure_tool_window') and self.measure_tool_window.winfo_exists():
            self.measure_tool_window.destroy()
        self.refocus()

    def measure_crystal(self,SeBaView_wrapper,function_to_run):
        self.crystal_size = [0,0]
        SeBaView_wrapper_rect = SeBaView_wrapper.rectangle()
        if hasattr(self,'measure_tool_window') and self.measure_tool_window.winfo_exists():
            self.measure_tool_window.destroy()
            self.measure_tool_window = Toplevel(self.root)
        else:
            self.measure_tool_window = Toplevel(self.root)
        icon = PhotoImage(file=icon_path)
        self.measure_tool_window.iconphoto(True,icon)
        self.measure_tool_window.title("Crystal Measuring Tool")
        self.measure_tool_window.geometry(f'{SeBaView_wrapper_rect.width()-self.screen_width // 4-10}x{SeBaView_wrapper_rect.height()}+{self.screen_width // 4-10}+{0}')
        self.measure_tool_window.resizable(FALSE,FALSE)
        self.measure_tool_window.attributes('-alpha','0.1')
        measure_tool = Canvas(self.measure_tool_window,width=self.measure_tool_window.winfo_width(),height=self.measure_tool_window.winfo_height(),bg='white')
        measure_tool.pack(fill='both',expand=True)
        self.mouse_pressed = False
        self.line_start = None
        self.line_end = None
        self.measure_tool_window.deiconify()
        self.measure_tool_window.lift()
        self.measure_tool_window.focus_force()

        def poll_mouse():
            nonlocal measure_tool
            global mouse_is_down
            if mouse_is_down and not self.mouse_pressed:
                self.mouse_pressed = True
                self.line_start = self.measure_tool_window.winfo_pointerxy()
            elif not mouse_is_down and self.mouse_pressed:
                self.mouse_pressed = False
                self.line_end = self.measure_tool_window.winfo_pointerxy()
                measure_tool.create_line(self.line_start[0]-self.screen_width // 4, self.line_start[1]-30,
                                            self.line_end[0]-self.screen_width // 4, self.line_end[1]-30, fill="blue", width=2)
                if self.crystal_size[0] == 0:
                    self.crystal_size[0] = int((((self.line_end[0] - self.line_start[0]) ** 2 +(self.line_end[1] - self.line_start[1]) ** 2) ** 0.5)*self.pixel_to_size)
                    print(f'crystal_width,crystal_height: {self.crystal_size[0]}, {self.crystal_size[1]}')
                    self.measure_tool_window.deiconify()
                    self.measure_tool_window.lift()
                    self.measure_tool_window.focus_force()
                elif self.crystal_size[0] != 0 and self.crystal_size[1] == 0:
                    self.crystal_size[1] = int((((self.line_end[0] - self.line_start[0]) ** 2+(self.line_end[1] - self.line_start[1]) ** 2) ** 0.5)*self.pixel_to_size)
                    print(f'crystal_width,crystal_height: {self.crystal_size[0]}, {self.crystal_size[1]}')
                    self.measure_tool_window.deiconify()
                    self.measure_tool_window.lift()
                    self.measure_tool_window.focus_force()
            if self.crystal_size[1] == 0:
                self.measure_tool_window.after(50,poll_mouse)
            else:
                if callable(function_to_run):
                    function_to_run()
        poll_mouse()

    def Harvest_Crystals(self):
        self.Select_Tray()

    def Upload_Xtal_Screen(self):
        self.clear_widgets()
        self.add_menu()
        upload_xtal_screen_frame = ttk.Frame(self.root,padding="3 3 12 12").grid(column=0,row=0,sticky=(N,W,E,S))

if __name__ == "__main__":
    app = CrystalDex_main()
    app.startup()
