#See the README for the functional goals of this program and for author notes and acknowledgements.
#https://byu.app.box.com/developers/console

#Imports
#General imports
import os
import glob
import shutil
import json
import openpyxl as px
from openpyxl.styles import PatternFill
from datetime import datetime
import time
import threading

#GUI imports
#https://tkdocs.com/tutorial/intro.html#audience
#https://tkdocs.com/tutorial/firstexample.html#design
import tkinter as tk
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
import pyperclip
from pywinauto.application import Application
import pyautogui
from pynput import mouse
import pywinauto.keyboard

#Crystal screen imports:
import pdfplumber

#Packaging stuff:
#https://realpython.com/pyinstaller-python/

#Paths
script_dir = os.path.dirname(os.path.abspath(__file__))#The directory of this script, so basically the folder where all the code is kept.
icon_path = os.path.join(script_dir,'Resources',"crystaldex_icon.png")
splash_path = os.path.join(script_dir,'Resources','CrystalDex_splash.png')
crystal_pictures = os.path.join(script_dir,'Resources',"Crystal_Pictures")
CrystalDex_library = os.path.join(script_dir,'Resources',"CrystalDex_Library.xlsx")
Crystal_Sendoff = os.path.join(script_dir,'Resources','Crystal_Sendoff_Sheet.xlsx')
crystal_screens_path = os.path.join(script_dir,'Resources','Crystal_Screens.json')
SeBaView_path = os.path.join(script_dir,'Resources','SeBaView_path_file.json')
os.makedirs(crystal_pictures,exist_ok=True)
home = os.path.expanduser("~")
downloads = os.path.join(home, "Downloads")
desktop = os.path.expanduser("~/Desktop")

#Dictionaries:
subwell_to_condition_dict = {
                f"{row}{col}": i
                for i,(row, col) in enumerate(
                    (r, c) for r in "ABCDEFGH" for c in range(1, 13)
                )
            }

well_to_excel_dict = {'A':8,'B':23,'C':38,'D':53,'E':68,'F':83,'G':98,'H':113,'1':'C','2':'H','3':'M','4':'R','5':'W','6':'AB','7':'AG','8':'AL','9':'AQ','10':'AV','11':'BA','12':'BF'}

#In the future, this can be used to allow new users to reconfigure the buttonpresses that are simulated on whatever microscope they're using.
def on_click(x,y,button,pressed):
    global mouse_is_down
    mouse_is_down = pressed

listener = mouse.Listener(on_click=on_click)
listener.start()

class CrystalDex_main:
    def __init__(self):
        self.chaperone_values = ["1TEL","2TEL","3TEL","4TEL","5TEL","6TEL"]
        self.tray_names = {}
        self.crystal_size = [0,0]
        self.harvesting = False
        self.editing = False
        self.pixel_to_size = 2000/867 #2 millimeter or 2000 microns per 867 pixels at 100% magnification (ie. a picture size of 1280x960pixels on the screen)
        self.picture_upload_filenames = {}
        self.button_location = None
        self.cell_fill_color = PatternFill(fill_type='solid',start_color='A9D18E',end_color='A9D18E')
        self.not_first = ''
        self.long_name = ''
        self.two_code = ''
        self.opened_microscope_app = False

        #Tkinter initializations
        root=tk.Tk()
        self.root = root
        self.root.title("CrystalDex")
        icon = tk.PhotoImage(file=icon_path)
        self.root.iconphoto(True,icon)
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        self.root.minsize(self.screen_width//5,600)
        self.root.geometry(f'1050x700+{self.screen_width//2-525}+{self.screen_height//2-350}')
        #Make the window resizable:
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)
        self.root.protocol("WM_DELETE_WINDOW", self.close_SeBaView_and_root)
        self.selected_condition = tk.StringVar(value='')

        #Box integrations:
        CLIENT_ID = "ywdxl21bfyxj6lpzest9alondci3jezf"
        CLIENT_SECRET = "WV4AhaJ4P0b6UHy8ENaXTNby6mjyxJv5"

        def delete_token_files(base_name = 'box_token'):
            for f in glob.glob(base_name+'*'):
                try:
                    os.remove(f)
                except Exception as e:
                    print(f'Failed to delete {f}: {e}')

        def authorize():
            token_storage = FileTokenStorage(filename='box_token')
            config = OAuthConfig(
                client_id=CLIENT_ID,
                client_secret=CLIENT_SECRET,
                token_storage=token_storage
            )
            return BoxOAuth(config)

        def get_code():
            for proc in reversed(list(psutil.process_iter(['pid','name']))):
                if proc.info['name'] and 'msedge.exe'.lower() in proc.info['name'].lower():
                    print(f'PID: {proc.info['pid']} - Title: {proc.info['name']}')#For some reason, deleting this line or commenting it out messes things up.
                    try:
                        browser = Application().connect(process=proc.info['pid'])
                        browser.top_window().set_focus()
                        start = time.time()
                        while time.time() - start< 120:
                            try:
                                pyautogui.hotkey('ctrl','l')
                                time.sleep(0.1)
                                pyautogui.hotkey('ctrl','c')
                                time.sleep(0.1)
                                url = pyperclip.paste()
                                if 'code=' in url and 'localhost' in url:
                                    #print(f'url: {url}')
                                    authorization_code = url.split('code=')[1]
                                    return authorization_code
                            except Exception as e:
                                print('Error while reading clipboard:', e)
                            time.sleep(10)
                        return True
                    except Exception:
                        pass
            return False

        auth = authorize()

        def reset_box_auth():
            print(f'Box token loading failed. Resetting token storage.')
            delete_token_files()
            auth = authorize()
            auth_url = auth.get_authorize_url()
            webbrowser.get('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe %s').open(auth_url)
            authorization_code = get_code()
            auth.get_tokens_authorization_code_grant(authorization_code)
            self.client = BoxClient(auth=auth)

        try:
            auth.retrieve_token()
            self.client = BoxClient(auth=auth)
            #print(f'User already approved app for Box.')
        except Exception as e:
            reset_box_auth()

        #Access CrystalDex_Library or the Mastercopy:
        file_download = None
        try:
            file_id = '1911115179608'
            try:
                file_download = self.client.downloads.download_file(file_id).read()
            except Exception:
                reset_box_auth()
                file_download = self.client.downloads.download_file(file_id).read()
        except FileNotFoundError:
                file_id = '1911117908557'
                file_download = self.client.downloads.download_file(file_id).read()
            
        #Begin writing a new CrystalDex_Library.xlsx file on the working computer:
        with open(CrystalDex_library,"wb") as c:
            c.write(file_download)
        self.wb = px.load_workbook(filename=os.path.abspath(CrystalDex_library))
        
        #Access Crystal_Sendoff_Sheet:
        file_download = None
        try:
            file_id = '1911117898957'
            file_download = self.client.downloads.download_file(file_id).read()
        except FileNotFoundError:
            file_id = '1911115194008'
            file_download = self.client.downloads.download_file(file_id).read()
        with open(Crystal_Sendoff,'wb') as s:
            s.write(file_download)
        self.sendoff_workbook = px.load_workbook(filename=os.path.abspath(Crystal_Sendoff))
        self.sendoff_sheet = self.sendoff_workbook['Crystal_Sendoff_Sheet']

        #Access Crystal_Screens.json:
        self.crystal_screens = {}
        file_download = None
        try:
            file_id = 1911117169207
            file_download = self.client.downloads.download_file(file_id).read()
        except FileNotFoundError:
            pass
        self.crystal_screen_values = []
        self.crystal_screen_symbols = {}
        if os.path.exists(crystal_screens_path):
            with open(crystal_screens_path, "r") as s:
                self.crystal_screens = json.load(s)
                for crystal_screen in self.crystal_screens.keys():
                    self.crystal_screen_values.append(crystal_screen)
                    self.crystal_screen_symbols[crystal_screen] = crystal_screen.split('__')[1]
        else:
            pass

    def reload_crystal_screens(self):
        if os.path.exists(crystal_screens_path):
            with open(crystal_screens_path, "r") as s:
                self.crystal_screens = json.load(s)

    def refocus(self):
        """Refocus the window if minimized"""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def Box_Save(self):
        """This method allows users to upload both pictures and workbooks into Box. I'm not sure if it will work for uploading to Box folders that are not within the user's 
        personal directory, but I will be testing that soon."""
        box_pics = self.client.folders.get_folder_items('328850048557', fields='name').entries
        box_pic_names = {box_pic.name for box_pic in box_pics}
        for image_filename in os.listdir(crystal_pictures):
            if image_filename in self.picture_upload_filenames.keys() and image_filename not in box_pic_names:
                file_path = os.path.join(crystal_pictures, image_filename)
                try:
                    with open(file_path,'rb') as image_stream:
                        uploading_file_return = self.client.uploads.upload_file(
                            UploadFileAttributes(
                                name=image_filename,parent=UploadFileAttributesParentField(id='328850048557') #The id here is where the images will end up. It references a folder in Box.
                            ),
                            image_stream
                        )
                        uploading_file = uploading_file_return.entries[0]
                        self.client.shared_links_files.add_share_link_to_file(file_id=uploading_file.id,fields='shared_link')
                        shared_link_id = self.client.shared_links_files.get_shared_link_for_file(uploading_file.id,"shared_link").id
                        shared_link_url = 'https://byu.app.box.com/file/'+str(shared_link_id)
                        ws_title = f'{self.picture_upload_filenames.get(image_filename)[0]}'
                        ws = self.wb[ws_title]
                        cell_id = self.picture_upload_filenames.get(image_filename)[1]
                        ws[cell_id].hyperlink = shared_link_url
                        self.wb.save(filename=os.path.abspath(CrystalDex_library))
                        if 'harvested' in image_filename:
                            for x in range(2,301):
                                cell_id = f'B{x}'
                                if self.sendoff_sheet[cell_id] == image_filename:
                                    self.sendoff_sheet[cell_id].hyperlink = shared_link_url
                        self.sendoff_workbook.save(filename=os.path.abspath(Crystal_Sendoff))
                    os.remove(file_path)
                    #Should self.picture_upload_filenames be reset at this point? If any of them gave an error, that would be bad (it would never upload). There doesn't seem to be an issue besides upload times being a little longer, so I'll leave it this way for now.
                except Exception as e:
                    print(f'Error uploading {image_filename}: {e}')

        #The following command uploads the Crystal Trays Library
        self.client.uploads.upload_file_version(
            attributes=box_sdk_gen.UploadFileAttributesParentField(
                name="CrystalDex_Library.xlsx",
                id="328850485682"),
                file_id="1911115179608",
                file=open(os.path.abspath(CrystalDex_library),"rb"
                )
            )
        
        #The following command uploads the Crystal_Screens.json
        self.client.uploads.upload_file_version(
            attributes=box_sdk_gen.UploadFileAttributesParentField(
                name="Crystal_Screens.json",
                id="328850485682"),
                file_id="1911117169207",
                file=open(os.path.abspath(crystal_screens_path),"rb"
                )
            )
        
        #The following command uploads the Crystal_Sendoff_Sheet.xlsx
        if self.harvesting:
            self.client.uploads.upload_file_version(
            attributes=box_sdk_gen.UploadFileAttributesParentField(
                name="Crystal_Sendoff_Sheet.xlsx",
                id="328850485682"),
                file_id="1911117898957",
                file=open(os.path.abspath(Crystal_Sendoff),"rb"
                )
            )

        if hasattr(self.splash_win,"winfo_exists") and self.splash_win.winfo_exists():
            self.root.after(0,self.splash_win.destroy)
        else:
            self.startup()

    def splash(self):
        self.splash_win = tk.Toplevel(self.root)
        self.splash_win.overrideredirect(True)
        self.splash_win.geometry(f'800x590+{self.screen_width//2-400}+{self.screen_height//2-510//2}')
        self.splash_image = tk.PhotoImage(file=splash_path)
        ttk.Label(self.splash_win,text='Loading CrystalDex',image=self.splash_image).pack(expand=True)
        self.splash_win.attributes('-topmost',True)

    def load_SeBaView(self):
        """This allows the user to open SeBaView software whenever CrystalDex is running. In the future, I'd like to add a configuration method that lets them choose other
        software and simulate the correct tk.Button presses, but that is currently beyond the scope of this project."""
        exe_path = None
        if os.path.exists(SeBaView_path):
            with open(SeBaView_path, "r") as s:
                exe_path = json.load(s).get("SeBaView_path")
        if not exe_path or not os.path.exists(exe_path):
            # Ask user to locate it if not found or invalid
            exe_path = filedialog.askopenfilename(
                title="Select the SeBaView executable",
                filetypes=[("Executable files", "*.exe")]
            )
        if exe_path:
            with open(SeBaView_path, "w") as s:
                json.dump({"SeBaView_path": exe_path}, s)
            self.SeBaView = Application(backend="uia").start(exe_path)
            time.sleep(4)
            try:
                SeBaView_main_window = self.SeBaView.window(title_re=".*SeBaView.*")
                self.SeBaView_wrapper = SeBaView_main_window.wrapper_object()
                self.SeBaView_wrapper.maximize()
                self.SeBaView_wrapper_rect = self.SeBaView_wrapper.rectangle()
                self.SeBaView_wrapper.set_focus()
                time.sleep(0.1)
                for i in range(2):
                    self.SeBaView_wrapper.click_input(coords=(60, 165))  #This accesses the camera connecting tk.Button.
                self.SeBaView_wrapper.minimize()
                self.opened_microscope_app = True
            except Exception as e:
                print("Failed to find or focus the SeBaView window. Restarting your computer usually fixes this issue.") #Change this to a Tkinter messagebox
                print(f"Error: {e}")
        self.root.after(0,self.splash_win.destroy)

    def close_SeBaView_and_root(self):
        try:
            self.SeBaView_wrapper.close()
        except:
            print(f'Failed to close SeBaView. Do it please!')
        self.root.destroy()

    def startup(self):
        self.harvesting = False
        self.editing = False
        if not self.opened_microscope_app:
            self.splash()
            threading.Thread(target=self.load_SeBaView,daemon=True).start()
        self.clear_widgets()
        self.add_menu()
        startup = ttk.Frame(self.root,padding='5 5 20 20')
        self.root.geometry(f'{self.screen_width}x{self.screen_height}+0+0')
        startup.option_add('*tearOFF',tk.FALSE)
        startup.grid(column=0,row=0,sticky='N,E,S,W')
        #To make the buttons bigger and prettier, you'll have to use another widget, probably a text widget with a tk.Button placed inside it.
        #https://tkdocs.com/tutorial/text.html#basics
        tk.Button(startup,text="Index Tray",command=self.New_Tray,width=40).grid(column=0,row=0,padx=50,pady=50,sticky='N,E,S,W')
        tk.Button(startup,text='Edit Tray',command=self.Edit_Tray,width=40).grid(column=1,row=0,padx=50,pady=50,sticky='N,E,S,W')
        tk.Button(startup,text='Harvest Crystals',command=self.Harvest_Crystals,width=40).grid(column=2,row=0,padx=50,pady=50,sticky='N,E,S,W')
        tk.Button(startup,text="Upload or Edit Crystal Screen",command=self.Upload_Crystal_Screen,width=40).grid(column=3,row=0,padx=50,pady=50,sticky='N,E,S,W')
        tk.Button(startup,text='Design and Upload Optimization Screen',command=self.Optimization_Screen,width=40).grid(column=0,row=1,padx=50,pady=50,sticky='N,E,S,W')

        for i in range(2):
            self.refocus()
        self.root.mainloop() #This has to be the last line of code in the startup function (it can't be placed anywhere before SeBaView is loaded)

    def add_menu(self):
        menu = tk.Menu(self.root)
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
        helpframe.grid(column=0,row=0,sticky='N,W,E,S')
        ttk.Label(helpframe,text="Welcome to CrystalDex, your helper for recording data from protein crystallization experiments!").grid(column=0,row=0,sticky='N,E,W')
        helptext = "This program functions by accessing Box and syncing with Excel sheets that contain links to every picture you take.\nCrystalDex allows you to run the microscope application within its GUI and prompts you to measure and label each crystal.\nIt then synchronizes all the crystallization screen data from its library of screens with each crystal picture taken.\nThere are other subprograms in this app that allow you to upload new crystallization screens into its library (such as for optimization screens). \nFor more assistance, reach out to miles.j.bradford@outlook.com"
        ttk.Label(helpframe,text=helptext).grid(column=0,row=1,sticky='N,E,W')
        self.refocus()

    def New_Tray(self):
        self.clear_widgets()
        self.add_menu()
        new_tray_frame = ttk.Frame(self.root,padding="3 3 12 12")
        new_tray_frame.grid(column=0,row=0,sticky='N,W')
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)

        ttk.Label(new_tray_frame, text="Select from standard tags or type a new entry:").grid(column=1,row=1)

        date_set_values = [str(datetime.now().strftime('%m-%d-%Y'))] #Replace with code that accesses a page in an excel workbook that contains the date_set_values of each tray in the CrystalDex.
        today_label = ttk.Label(new_tray_frame,text="Today?")
        today_label.grid(column=3,row=5,sticky='N,W')
        today_var = tk.BooleanVar()
        ttk.Checkbutton(new_tray_frame,variable=today_var,onvalue=True,offvalue=False).grid(column=4,row=5,sticky='N,W')
        date_set_label = ttk.Label(new_tray_frame,text="Date Set (required; 00-00-0000):")
        date_set_label.grid(column=1,row=5,sticky='N,W')
        date_set_var = tk.StringVar()
        date_set_drop_down = ttk.Combobox(new_tray_frame,textvariable=date_set_var,values=date_set_values)
        date_set_drop_down.grid(column=2,row=5)

        chaperone_label = ttk.Label(new_tray_frame,text="Crystal Chaperone (optional):")
        chaperone_label.grid(column=1,row=6,sticky='N,W')
        chaperone_var = tk.StringVar()
        chaperone_drop_down = ttk.Combobox(new_tray_frame,textvariable=chaperone_var,values=self.chaperone_values)
        chaperone_drop_down.grid(column=2,row=6)

        crystal_screen_label = ttk.Label(new_tray_frame,text="Crystal Screen (required):")
        crystal_screen_label.grid(column=1,row=7,sticky='N,W')
        crystal_screen_var = tk.StringVar()
        crystal_screen_drop_down = ttk.Combobox(new_tray_frame,textvariable=crystal_screen_var,values=self.crystal_screen_values)
        crystal_screen_drop_down.grid(column=2,row=7)

        target_protein_values = ["DARPin","CMG2","UBA","TELSAM","sfGFP"]
        target_protein_label = ttk.Label(new_tray_frame,text="Target Protein (For large constructs, include the full name with linkers):")
        target_protein_label.grid(column=1,row=8,sticky='N,W')
        target_protein_var = tk.StringVar()
        target_protein_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_var,values=target_protein_values)
        target_protein_drop_down.grid(column=2,row=8,sticky='N,W')

        target_protein_stock_concentration_values = ['1','5','15','20']
        target_protein_top_left_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into top left subwell (required):")
        target_protein_top_left_stock_concentration_label.grid(column=1,row=9,sticky='N,W')
        target_protein_top_left_stock_concentration_var = tk.StringVar()
        target_protein_top_left_stock_concentration_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_top_left_stock_concentration_var,values=target_protein_stock_concentration_values)
        target_protein_top_left_stock_concentration_drop_down.grid(column=2,row=9,sticky='N,W')

        target_protein_top_right_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into top right subwell (required):")
        target_protein_top_right_stock_concentration_label.grid(column=1,row=10,sticky='N,W')
        target_protein_top_right_stock_concentration_var = tk.StringVar()
        target_protein_top_right_stock_concentration_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_top_right_stock_concentration_var,values=target_protein_stock_concentration_values)
        target_protein_top_right_stock_concentration_drop_down.grid(column=2,row=10,sticky='N,W')

        target_protein_bottom_left_stock_concentration_label = ttk.Label(new_tray_frame,text="Target protein stock concentration placed into bottom left subwell (required):")
        target_protein_bottom_left_stock_concentration_label.grid(column=1,row=11,sticky='N,W')
        target_protein_bottom_left_stock_concentration_var = tk.StringVar()
        target_protein_bottom_left_stock_concentration_drop_down = ttk.Combobox(new_tray_frame,textvariable=target_protein_bottom_left_stock_concentration_var,values=target_protein_stock_concentration_values)
        target_protein_bottom_left_stock_concentration_drop_down.grid(column=2,row=11,sticky='N,W')

        #Later, if I have time, I'll want to add a little virtual replica in column 3 of a single well (with the four subwells) so that the user can see exactly what they're filling out, and each subwell will have the concentration appear as they fill it in.

        custom_tags_values = []
        custom_tags_label = ttk.Label(new_tray_frame,text="Custom Tags (optional; separated by commas, please!):")
        custom_tags_label.grid(column=1,row=12,sticky='N,W')
        custom_tags_var = tk.StringVar()
        custom_tags_drop_down = ttk.Combobox(new_tray_frame,textvariable=custom_tags_var,values=custom_tags_values)
        custom_tags_drop_down.grid(column=2,row=12)

        tk.Button(new_tray_frame,text="Begin Indexing Tray",
                   command=lambda: self.Index_Tray(
                       str(date_set_var.get()),
                       bool(today_var.get()),
                       str(crystal_screen_var.get()),
                       str(target_protein_var.get()),
                       str(target_protein_top_left_stock_concentration_var.get()),
                       str(target_protein_top_right_stock_concentration_var.get()),
                       str(target_protein_bottom_left_stock_concentration_var.get()),
                       str(chaperone_var.get()),
                       str(custom_tags_var.get())
                        )
                    ).grid(column=1,row=13,sticky='N,W')

        for child in new_tray_frame.winfo_children():
            child.grid_configure(padx=5,pady=5)
        self.refocus()

    def Select_Tray(self,short_title=None):
        self.clear_widgets()
        self.add_menu()
        st_frame = ttk.Frame(self.root,padding="3 3 12 12")
        st_frame.grid(column=0,row=0,sticky='N,W')
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)
        if not self.editing:
            st_name_label = ttk.Label(st_frame,text=(
                'At least one previously indexed tray was found that shares a date, screen, and target protein with the current tray.'
                '\nPlease review the following to ensure no duplicate trays are indexed!'
            ))
            st_name_label.grid(column=0,row=0)
        else:
            st_name_label = ttk.Label(st_frame,text=('Please select a tray to edit.'))
            st_name_label.grid(column=0,row=0)
        tray_name = tk.StringVar()
        st_name_combobox = ttk.Combobox(st_frame,textvariable=tray_name,values=list(self.tray_names.keys()))
        st_name_combobox.grid(column=0,row=1)

        if not self.harvesting and not self.editing:
            none_of_the_above_label = ttk.Label(st_frame,text="If none of the above match your tray, click 'make new tray':")
            none_of_the_above_label.grid(column=0,row=2)
            tk.Button(st_frame,text="make new tray",command=lambda: make_new_tray(short_title)).grid(column=1,row=2,sticky='W')

        def make_new_tray(short_title):
            shorter_title = short_title
            suffix = 1
            all_titles = [ws.title for ws in self.wb]

            # Make sure the title is unique
            while shorter_title in all_titles:
                shorter_title = f"{short_title[:24]}_{suffix}"
                suffix += 1

            new_worksheet = self.wb.copy_worksheet(self.wb["Mastercopy"])
            new_worksheet.title = shorter_title
            self.proceed(self.wb[shorter_title])

        tk.Button(st_frame,text="Save selection and proceed",
        command=lambda: self.proceed(self.wb[self.tray_names.get(tray_name.get())])).grid(column=0,row=3)
        self.refocus()

    def proceed(self,ws):
        date_set = ws['D1'].value
        chaperone = ws['D2'].value
        crystal_screen = ws['D3'].value
        target_protein = ws['D4'].value
        target_protein_top_left_stock_concentration = ws['H1'].value
        target_protein_top_right_stock_concentration = ws['H2'].value
        target_protein_bottom_left_stock_concentration = ws['H3'].value
        custom_tags_values = ws['D5'].value
        self.SeBaView_wrapper.maximize()
        self.SeBaView_wrapper.set_focus()
        self.identify_subwell(ws,date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,chaperone,custom_tags_values)
        self.refocus()

    def Index_Tray(self,date_set,today,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,chaperone,custom_tags_values):
        indexable = False
        self.SeBaView_wrapper.maximize()
        self.SeBaView_wrapper.set_focus()
        if today:
            date_set = (datetime.now().strftime('%m-%d-%Y'))
        else:
            try:
                date_set = str(datetime.strftime(datetime.strptime(date_set,"%m-%d-%Y"),"%m-%d-%Y"))
                indexable = True
            except ValueError:
                messagebox.showerror(title="Date Error",message="You attempted to put in an invalid date. Please use the style: 01-01-2025")
        if crystal_screen in self.crystal_screens.keys():
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
                if all(term in tags for term in [date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,chaperone]):
                    self.tray_names[str(ws['A1'].value)] = ws.title
                    ws_possible_duplicate_count += 1
            if ws_possible_duplicate_count >0:
                full_title = f'{date_set}_{self.crystal_screen_symbols.get(crystal_screen)}_{target_protein}_1'
                short_title = full_title[:26]
                self.Select_Tray(short_title)
            elif ws_possible_duplicate_count == 0:
                print(f"No trays found with these stats; generating new tray!")#change this to a Tkinter messagebox or the splash screen
                new_worksheet = self.wb.copy_worksheet(self.wb["Mastercopy"])
                full_title = f'{date_set}_{self.crystal_screen_symbols.get(crystal_screen)}_{target_protein}_1'
                short_title = full_title[:26]
                new_worksheet.title = short_title
                all_tags = [date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,chaperone,custom_tags_values]
                new_worksheet['A1'] = full_title
                new_worksheet['K1'] = ', '.join(map(str,all_tags))
                new_worksheet['D1'] = date_set
                new_worksheet['D2'] = chaperone
                new_worksheet['D3'] = crystal_screen
                new_worksheet['D4'] = target_protein
                new_worksheet['D5'] = custom_tags_values
                new_worksheet['H1'] = target_protein_top_left_stock_concentration
                new_worksheet['H2'] = target_protein_top_right_stock_concentration
                new_worksheet['H3'] = target_protein_bottom_left_stock_concentration
                self.wb.save(filename=os.path.abspath(CrystalDex_library))
                self.identify_subwell(new_worksheet,date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,chaperone,custom_tags_values)
        self.refocus()

    def identify_subwell(self,ws,date_set,crystal_screen,target_protein,target_protein_top_left_stock_concentration,target_protein_top_right_stock_concentration,target_protein_bottom_left_stock_concentration,chaperone,custom_tags_values):
        self.clear_widgets()
        self.add_menu()
        self.root.geometry(f"{self.screen_width // 4}x{self.screen_height}+0+0")
        subwell_frame = ttk.Frame(self.root,padding="3 3 12 12")
        subwell_frame.grid(column=0,row=0,sticky='N,W')
        self.root.columnconfigure(0,weight=1)
        self.root.rowconfigure(0,weight=1)
        self.refocus()

        ensure_magnified_label = ttk.Label(subwell_frame,text="MAKE SURE the microscope is fully\nmagnified before taking any pictures.\nALSO ENSURE that the SeBaView\n camera is at 80% magnification.")
        ensure_magnified_label.grid(column=1,row=1)

        well_row_label = ttk.Label(subwell_frame,text="Well row:")
        well_row_label.grid(column=1,row=2)
        well_row_values = ['A','B','C','D','E','F','G','H']
        well_row_var = tk.StringVar()
        well_row_drop_down = ttk.Combobox(subwell_frame,textvariable=well_row_var,values=well_row_values)
        well_row_drop_down.grid(column=2,row=2)

        well_column_label = ttk.Label(subwell_frame,text="Well column:")
        well_column_label.grid(column=1,row=3)
        well_column_values = ['1','2','3','4','5','6','7','8','9','10','11','12']
        well_column_var = tk.StringVar()
        well_column_drop_down = ttk.Combobox(subwell_frame,textvariable=well_column_var,values=well_column_values)
        well_column_drop_down.grid(column=2,row=3)

        subwell_values = ['top_left','top_right','bottom_left']
        subwell_label = ttk.Label(subwell_frame,text="subwell:")
        subwell_label.grid(column=1,row=4)
        subwell_var = tk.StringVar()
        subwell_drop_down = ttk.Combobox(subwell_frame,textvariable=subwell_var,values=subwell_values)
        subwell_drop_down.grid(column=2,row=4)

        crystal_width_var = tk.StringVar()
        crystal_width_label = ttk.Label(subwell_frame,text='crystal width:')
        crystal_width_label.grid(column=1,row=5)
        crystal_width_entry = ttk.Entry(subwell_frame,textvariable=crystal_width_var,state=tk.DISABLED)
        crystal_width_entry.grid(column=2,row=5)
        um_width_label = ttk.Label(subwell_frame,text='um')
        um_width_label.grid(column=3,row=5)

        crystal_height_var = tk.StringVar()
        crystal_height_label = ttk.Label(subwell_frame,text='crystal height:')
        crystal_height_label.grid(column=1,row=6)
        crystal_height_entry = ttk.Entry(subwell_frame,textvariable=crystal_height_var,state=tk.DISABLED)
        crystal_height_entry.grid(column=2,row=6)
        um_row_label = ttk.Label(subwell_frame,text='um')
        um_row_label.grid(column=3,row=6)

        number_of_crystals_label = ttk.Label(subwell_frame,text='# of harvestable crystals (optional):')
        number_of_crystals_label.grid(column=1,row=7)
        number_of_crystals_var = tk.StringVar()
        number_of_crystals_entry = ttk.Spinbox(subwell_frame,from_=0,to=100,textvariable=number_of_crystals_var)
        number_of_crystals_entry.grid(column=2,row=7)

        shape_label = ttk.Label(subwell_frame,text='Shape of crystals:')
        shape_label.grid(column=1,row=8)
        shape_var = tk.StringVar()
        shape_entry = ttk.Entry(subwell_frame,textvariable=shape_var)
        shape_entry.grid(column=2,row=8)

        possible_salt_crystals_label = ttk.Label(subwell_frame,text="Possibly a salt crystal")
        possible_salt_crystals_label.grid(column=1,row=9)
        possible_salt_crystals_var = tk.BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=possible_salt_crystals_var,onvalue=True,offvalue=False).grid(column=2,row=9)

        precipitation_label = ttk.Label(subwell_frame,text="Precipitation present")
        precipitation_label.grid(column=1,row=10)
        precipitation_var = tk.BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=precipitation_var,onvalue=True,offvalue=False).grid(column=2,row=10)

        microcrystals_label = ttk.Label(subwell_frame,text="Microcrystals present")
        microcrystals_label.grid(column=1,row=11)
        microcrystals_var = tk.BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=microcrystals_var,onvalue=True,offvalue=False).grid(column=2,row=11)

        glassy_protein_or_artifacts_label = ttk.Label(subwell_frame,text="Glassy protein or artifacts present")
        glassy_protein_or_artifacts_label.grid(column=1,row=12)
        glassy_protein_or_artifacts_var = tk.BooleanVar()
        ttk.Checkbutton(subwell_frame,variable=glassy_protein_or_artifacts_var,onvalue=True,offvalue=False).grid(column=2,row=12)

        now = datetime.now()
        date_snapped = now.strftime('%m-%d-%Y-%H-%M-%S')

        x = 0
        if self.harvesting:
            x = 2
            harvester_label = ttk.Label(subwell_frame,text='Full name of harvester:')
            harvester_label.grid(column=1,row=13)
            harvester_var = tk.StringVar()
            harvester_entry = ttk.Entry(subwell_frame,textvariable=harvester_var)
            harvester_entry.grid(column=2,row=13)

            vials_available = [range(2,301)]
            for y in range(2,301):
                cell_id = f'B{y}'
                if self.sendoff_sheet[cell_id].value != None:
                    vials_available.remove(y)
            vial_label = ttk.Label(subwell_frame,text='Enter vial number:')
            vial_label.grid(column=1,row=14)
            vial_var = tk.StringVar()
            vial_dropdown = ttk.Combobox(subwell_frame,textvariable=vial_var,values=vials_available)
            vial_dropdown.grid(column=2,row=14)

        notes_label = ttk.Label(subwell_frame,text="Crystallographer notes:")
        notes_label.grid(column=1,row=13+x)
        notes = tk.Text(subwell_frame, width = 50, height = 5)
        notes.grid(column=1,row=14+x,columnspan=2)

        def update_crystal_size_vars():
            crystal_width_var.set(f'{self.crystal_size[0]}')
            crystal_height_var.set(f'{self.crystal_size [1]}')

        tk.Button(subwell_frame,text ='Measure Crystal',
                   command=lambda: self.measure_crystal(update_crystal_size_vars)).grid(column=1,row=15+x)

        if self.harvesting:
            tk.Button(subwell_frame,text='Harvest crystal',
                    command=lambda: self.take_picture(
                        ws,
                        chaperone,
                        target_protein,
                        crystal_screen,
                        str(well_row_var.get()),
                        str(well_column_var.get()),
                        str(subwell_var.get()),
                        str(number_of_crystals_var.get()),
                        str(shape_var.get()),
                        bool(possible_salt_crystals_var.get()),
                        bool(precipitation_var.get()),
                        bool(microcrystals_var.get()),
                        bool(glassy_protein_or_artifacts_var.get()),
                        date_set,
                        date_snapped,
                        notes.get(1.0,tk.END),
                        vial=str(vial_var.get()),
                        harvester = str(harvester_var.get())
                    )).grid(column=1,row=16+x)
        else:
            tk.Button(subwell_frame,text='Take and save picture',
                    command=lambda: self.take_picture(
                        ws,
                        chaperone,
                        target_protein,
                        crystal_screen,
                        str(well_row_var.get()),
                        str(well_column_var.get()),
                        str(subwell_var.get()),
                        str(number_of_crystals_var.get()),
                        str(shape_var.get()),
                        bool(possible_salt_crystals_var.get()),
                        bool(precipitation_var.get()),
                        bool(microcrystals_var.get()),
                        bool(glassy_protein_or_artifacts_var.get()),
                        date_set,
                        date_snapped,
                        notes.get(1.0,tk.END)
                    )).grid(column=1,row=16+x)

        tk.Button(subwell_frame,text="Done with this tray",
                   command=lambda: self.Box_Save()).grid(column=1,row=17+x)
        
        for child in subwell_frame.winfo_children():
            child.grid_configure(padx=5,pady=10)
        
    def monitor_mouse(self):
        print(f'monitor_mouse is running...')
        if mouse_is_down:
            x, y = pyautogui.position()
            rel_x = x - self.SeBaView_wrapper_rect.left
            rel_y = y - self.SeBaView_wrapper_rect.top
            print(f'Mouse pressed relative to SeBaView window: ({rel_x}, {rel_y})')
            self.button_location = rel_x,rel_y
        time.sleep(0.1)
        print(f'monitor mouse is done.')

    def take_picture(
            self,
            ws,
            chaperone,
            target_protein,
            crystal_screen,
            well_row,
            well_column,
            subwell,
            number_of_crystals,
            shape,
            possible_salt_crystals,
            precipitation,
            microcrystals,
            glassy_protein_or_artifacts,
            date_set,
            date_snapped,
            notes,
            vial=None,
            harvester=None):
        """This is the pride and jewel of CrystalDex, which allows users to take a picture, name it, and upload it all at once without any extra hassle."""
        
        image_title = f'{chaperone}_{target_protein}_{crystal_screen}_{well_row}{well_column}_{subwell}_{date_set}_{date_snapped}'
        if self.harvesting:
            image_title = image_title+'_harvested'

        def take_take_picture():
            self.SeBaView_wrapper.maximize()
            self.SeBaView_wrapper.set_focus()
            self.SeBaView_wrapper.click_input(coords=(55, 70))  #This accesses the save as tk.Button.
            time.sleep(1)
            for i in range(2):
                self.SeBaView_wrapper.click_input(coords=(750,450))#This is supposed to access the Desktop tk.Button to save the photos there temporarily, although it might not work. I may need to get a wrapper for the save window that opens... 
            pywinauto.keyboard.send_keys(f"{image_title}{{ENTER}}") #Enter the image_title name into the save window
            row = well_to_excel_dict.get(well_row)
            column = well_to_excel_dict.get(well_column)
            
            picture_link_cell = ws[f'{column}{row}']
            if subwell=='top_right':
                picture_link_cell = picture_link_cell.offset(row=0,column=2)
                column = px.utils.get_column_letter(px.utils.column_index_from_string(column)+2)
            elif subwell=='bottom_left':
                picture_link_cell = picture_link_cell.offset(row=7,column=0)
                row = row+7
            picture_link_cell.value = image_title
            picture_link_cell.offset(row=1,column=0).value = f'{(datetime.strptime(date_snapped,'%m-%d-%Y-%H-%M-%S')-datetime.strptime(date_set,"%m-%d-%Y")).days}' #might have to change the type of these variables
            picture_link_cell.offset(row=2,column=0).value = f'{self.crystal_size[0]}x{self.crystal_size[1]} um'
            picture_link_cell.offset(row=3,column=0).value = f'{number_of_crystals}'
            picture_link_cell.offset(row=4,column=0).value = f'{shape}'
            picture_link_cell.offset(row=5,column=0).value = f'Possible salt crystals: {possible_salt_crystals}, precipitation: {precipitation}, microcrystals: {microcrystals}, glassy protein or artifacts: {glassy_protein_or_artifacts}'
            picture_link_cell.offset(row=6,column=0).value = notes

            for x in range(7):
                picture_link_cell.offset(row=x,column=-1).fill = self.cell_fill_color
                picture_link_cell.offset(row=x,column=0).fill = self.cell_fill_color
                
            self.wb.save(filename=os.path.abspath(CrystalDex_library))

            for filename in os.listdir(desktop):
                file_path = os.path.join(desktop, filename)
                if os.path.isfile(file_path) and filename.lower().endswith(('.jpeg','.jpg','.bmp')) :#and time.time() - os.path.getmtime(file_path)<10
                    try:
                        shutil.move(file_path, crystal_pictures)
                        self.picture_upload_filenames[filename] = [ws.title,f'{column}{row}']
                    except Exception as e:
                        print(f"Failed to move {filename}: {e}")
                        print(f'Still placing filename within self.picture_upload_filenames to be uploaded.')
                        self.picture_upload_filenames[filename] = [ws.title,f'{column}{row}']

            if hasattr(self,'measure_tool_window') and self.measure_tool_window.winfo_exists():
                self.measure_tool_window.destroy()
            
        if self.harvesting == False:
            take_take_picture()
        elif self.harvesting:
            if self.crystal_size[1] != 0:
                take_take_picture()
                #The following updates the Crystal_Sendoff_Sheet and does so every time a harvested picture is taken. This is not uploaded to Box until the Box_Save function is run when you're done with the whole tray.
                cell_id = f'B{vial}'
                crystal_cell = self.sendoff_sheet[cell_id]
                crystal_cell.value = image_title
                condition = self.crystal_screens.get(crystal_screen)[subwell_to_condition_dict[f'{well_row}{well_column}']]
                crystal_cell.offset(row=0,column=1).value = condition
                crystal_cell.offset(row=0,column=2).value = f'{shape}'
                crystal_cell.offset(row=0,column=3).value = min(self.crystal_size[0],self.crystal_size[1]) #Minor axis
                crystal_cell.offset(row=0,column=4).value = max(self.crystal_size[0],self.crystal_size[1]) #Major axis
                crystal_cell.offset(row=0,column=5).value = harvester
                crystal_cell.offset(row=0,column=6).value = f'{date_set}, {date_snapped}' #date_set and date_harvested are passed from identify_subwell 
                crystal_cell.offset(row=0,column=7).value = notes
                self.sendoff_workbook.save(filename=os.path.abspath(Crystal_Sendoff))
            else:
                messagebox.showerror(title='No crystal measurement',message="You haven't measured your crystal, silly!")

        self.refocus()

    def measure_crystal(self,function_to_run):
        """This is one of the best features of CrystalDex! However, it does need a calibrate tk.Button. Currently, it only works for the microscope in Dr. Moody's lab at BYU.
        """
        self.crystal_size = [0,0]
        if hasattr(self,'measure_tool_window') and self.measure_tool_window.winfo_exists():
            self.measure_tool_window.destroy()
            self.measure_tool_window = tk.Toplevel(self.root)
        else:
            self.measure_tool_window = tk.Toplevel(self.root)
        icon = tk.PhotoImage(file=icon_path)
        self.measure_tool_window.iconphoto(True,icon)
        self.measure_tool_window.title("Crystal Measuring Tool")
        self.SeBaView_wrapper_rect = self.SeBaView_wrapper.rectangle()
        self.measure_tool_window.geometry(f'{self.SeBaView_wrapper_rect.width()-self.screen_width//4-10}x{self.SeBaView_wrapper_rect.height()}+{self.screen_width // 4-10}+{0}')
        self.measure_tool_window.resizable(tk.FALSE,tk.FALSE)
        self.measure_tool_window.attributes('-alpha','0.1')
        measure_tool = tk.Canvas(self.measure_tool_window,width=self.measure_tool_window.winfo_width(),height=self.measure_tool_window.winfo_height(),bg='white')
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
                    self.measure_tool_window.deiconify()
                    self.measure_tool_window.lift()
                    self.measure_tool_window.focus_force()
                elif self.crystal_size[0] != 0 and self.crystal_size[1] == 0:
                    self.crystal_size[1] = int((((self.line_end[0] - self.line_start[0]) ** 2+(self.line_end[1] - self.line_start[1]) ** 2) ** 0.5)*self.pixel_to_size)
                    self.measure_tool_window.deiconify()
                    self.measure_tool_window.lift()
                    self.measure_tool_window.focus_force()
            if self.crystal_size[1] == 0:
                self.measure_tool_window.after(50,poll_mouse)
            else:
                if callable(function_to_run):
                    function_to_run()
        poll_mouse()

    def Edit_Tray(self):
        """This is the root method that allows you to index a tray that has already been started."""
        self.editing = True
        x = 0
        for ws in self.wb:
            x += 1
            self.tray_names[str(ws['A1'].value)] = ws.title
        if x>1:
            self.Select_Tray()
        else:
            messagebox.showerror(title='No crystal trays indexed yet...',message=f"There are no trays in your CrystalDex Library. You can't harvest what doesn't exist!")
            self.startup()

    def Harvest_Crystals(self):
        """This is the root method that allows you to use the Select_Tray method in the harvesting mode, which forces you to include some types of data, but which also
        creates a crystal sendoff sheet that is useful for tracking what conditions and stats each crystal had.
        """
        self.harvesting = True
        x = 0
        for ws in self.wb:
            x += 1
            self.tray_names[str(ws['A1'].value)] = ws.title
        if x>1:
            self.Select_Tray()
        else:
            messagebox.showerror(title='No crystal trays indexed yet...',message=f"There are no trays in your CrystalDex Library. You can't harvest what doesn't exist!")
            self.startup()

    def Upload_Crystal_Screen(self):
        """This method allows users to upload a crystal screen directly from Hampton's data sheets. It doesn't always work, but it does luckily have a method for overwriting
        the results before you save the crystal screen. NOTE: There is currently no way for users to delete crystal screens. I may need to add this later.
        """
        self.reload_crystal_screens()
        self.clear_widgets()
        self.add_menu()
        self.root.geometry(f'1250x700+{self.screen_width//2-625}+{self.screen_height//2-350}')
        upload_crystal_screen_frame = ttk.Frame(self.root,padding="3 3 12 12")
        upload_crystal_screen_frame.grid(column=0,row=0,sticky='N,W,E,S')
        
        crystal_screen_name_label = ttk.Label(upload_crystal_screen_frame,text='Enter the name of the new crystal screen:')
        crystal_screen_name_label.grid(row=0,column=0)
        crystal_screen_name = tk.StringVar()
        crystal_screen_entry = ttk.Entry(upload_crystal_screen_frame,textvariable=crystal_screen_name)
        crystal_screen_entry.grid(row=0,column=1)

        crystal_screen_symbol_label = ttk.Label(upload_crystal_screen_frame,text='Enter 2-letter symbol for new screen:')
        crystal_screen_symbol_label.grid(row=0,column=2)
        crystal_screen_symbol = tk.StringVar()
        crystal_screen_symbol_entry = ttk.Entry(upload_crystal_screen_frame,textvariable=crystal_screen_symbol)
        crystal_screen_symbol_entry.grid(row=0,column=3)

        upload_crystal_screen_button = tk.Button(upload_crystal_screen_frame,text=f"Upload {self.not_first}crystal screen",command=lambda: scrape_crystal_screen_data())
        upload_crystal_screen_button.grid(column=4,row=0,sticky='N,W')
        upload_crystal_screen_button.configure(text=f'Upload {self.not_first}crystal screen')

        conditions = ['' for _ in range(96)]

        def scrape_crystal_screen_data():
            crystal_screen_path = filedialog.askopenfilename(
                    title="Select the pdf containing the crystal screen conditions",
                    filetypes=[('PDF files',"*.pdf")]
                )
            text = ""
            with pdfplumber.open(crystal_screen_path) as pdf:
                for page in pdf.pages:
                    text += page.extract_text() + "\n"

                if not any(keyword in text for keyword in ['%','magnesium','calcium','chloride','cobalt','nickel','polyethylene','glycol','monomethyl','ether','tris','none','potassium','sodium','tartrate','tetrahydrate','trihydrate','hydrochloride','hexahydrate','dihydrate','ammonium']):
                    print(f"No text found. If you're trying to use Make Tray from Hampton, use the Optimization tk.Button on the home screen instead.")

            last_condition = 0
            offset = 0
            for condition in range(96):
                if conditions[condition].strip():
                    last_condition = condition
                    offset = 1

            for condition in range(96):
                next_condition = str(condition+2)
                reading = False
                for line in text.splitlines():
                    if line.startswith(f'{condition+1}.'):
                        line = line.strip()
                        if (any(keyword in line.lower() for keyword in ['%','magnesium','calcium','chloride','cobalt','nickel','polyethylene','glycol','monomethyl','ether','tris','none','potassium','sodium','tartrate','tetrahydrate','trihydrate','hydrochloride','hexahydrate','dihydrate','ammonium'])) or all(keyword in line.lower() for keyword in ['ph',' m ']):
                            try:
                                conditions[condition+last_condition+offset] = line
                                reading = True
                            except IndexError:
                                messagebox.showerror(title='Too many conditions',message=f'There were too many conditions to add to the new screen. Please review the upload.')
                                break
                                #Note that this is a very rough fix to the problem.
                    elif reading and not line.startswith(f'{next_condition}.'):
                            if any(keyword in line.lower() for keyword in ['ide','ate','magnesium','calcium','chloride','cobalt','nickel','polyethylene','glycol','monomethyl','ether','tris','none','potassium','sodium','tartrate','tetrahydrate','trihydrate','hydrochloride','hexahydrate','dihydrate','ammonium']):
                                conditions[condition+last_condition+offset] += ' ' + line
                    elif reading and line.startswith(f'{next_condition}.'):
                        reading = False

            listbox_label = ttk.Label(upload_crystal_screen_frame,text='Review and correct generated conditions:')
            listbox_label.grid(row=2,column=0)
            listbox_values = [f"[{condition+1}]: {conditions[condition]}" for condition in range(len(conditions))]
            condition_var = tk.StringVar(value=listbox_values)
            conditions_listbox = tk.Listbox(upload_crystal_screen_frame,listvariable=condition_var,height=25,width=150)
            conditions_listbox.grid(row=3,column=0,columnspan=3)
            
            edited_condition = tk.StringVar()
            condition_entry = tk.Entry(upload_crystal_screen_frame, textvariable=edited_condition, width=150)
            condition_entry.grid(row=4, column=0, columnspan=3)

            selected_index = tk.IntVar(value=-1)

            def select_condition(event):
                selection = conditions_listbox.curselection()
                if selection:
                    index = selection[0]
                    selected_index.set(index)
                    edited_condition.set(f'{conditions[index]}')

            conditions_listbox.bind('<<ListboxSelect>>',select_condition)

            def overwrite():
                index = selected_index.get()
                print(f'index: {index}')
                if index >= 0:
                    text = edited_condition.get()
                    condition = index
                    conditions[condition] = text
                    conditions_listbox.delete(index)
                    conditions_listbox.insert(index, f'[{condition+1}]: {text}')
                    selected_index.set(index+1)
                    if index !=95:
                        edited_condition.set(f'{conditions[index+1]}')
                    else:
                        edited_condition.set(f'{conditions[0]}')
                    
            tk.Button(upload_crystal_screen_frame,text='overwrite',command=overwrite).grid(row=4,column=3)
            
            def save():
                self.crystal_screens[f'{crystal_screen_name.get()}__{crystal_screen_symbol.get()}'] = conditions
                with open(crystal_screens_path, "w") as c:
                    json.dump(self.crystal_screens, c)
                self.close_SeBaView_and_root()#For some reason, the json won't upload until after the tkinter root is closed.
                self.Box_Save()
            
            tk.Button(upload_crystal_screen_frame,text='Save and finish',command=save).grid(row=5,column=2)

    def Optimization_Screen(self):
        """Allows users to either create a custom optimization screen (while optionally looking up a reference condition from any of the screens
        currently in CrystalDex) or to copy information (by hand) from a reference sheet made by Hampton's Make Tray. NOTE: This will not work 
        for any Make Tray optimizations that have conditions that are optimized in a non-linear manner."""
        self.clear_widgets()
        self.add_menu()
        self.root.geometry(f'1250x700+{self.screen_width//2-625}+{self.screen_height//2-350}')
        optimization_screen_frame = ttk.Frame(self.root,padding="3 3 12 12")
        optimization_screen_frame.grid(column=0,row=0,sticky='N,W,E,S')

        self.reload_crystal_screens()
        self.crystal_screen = None
        conditions = ['' for _ in range(96)]
        selected_index = tk.IntVar(value=-1)
        self.quad = 1

        ttk.Label(optimization_screen_frame,text='Fill out the following to name your optimization screen. Be advised that CrystalDex appends the date to each optimization screen as\n' \
        'this is often one of the most defining characteristics of any tray/screen and helps to avoid duplicate names.',justify='left').grid(column=0,row=0,columnspan=2)

        ttk.Label(optimization_screen_frame,text='Complete name of new optimization screen:').grid(row=1,column=0)
        long_name = tk.StringVar()
        long_name_entry = tk.Entry(optimization_screen_frame,textvariable=long_name)
        long_name_entry.grid(row=1,column=1)

        ttk.Label(optimization_screen_frame,text='Two-character code for optimization screen:').grid(row=2,column=0)
        two_code = tk.StringVar()
        two_code_entry = tk.Entry(optimization_screen_frame,textvariable=two_code)
        two_code_entry.grid(row=2,column=1)

        tk.Button(optimization_screen_frame,text='Continue',command=lambda: select_type(long_name_entry.get(),two_code_entry.get())).grid(row=3,column=0)

        def select_type(long_name,two_code):
            self.long_name = long_name
            self.two_code = two_code
            self.clear_widgets()
            self.add_menu()
            self.root.geometry(f'1250x700+{self.screen_width//2-625}+{self.screen_height//2-350}')
            optimization_screen_frame = ttk.Frame(self.root,padding="3 3 12 12")
            optimization_screen_frame.grid(column=0,row=0,sticky='N,W,E,S')

            ttk.Label(optimization_screen_frame,text='Use Make Tray by Hampton and enter your desired condition(s) to generate a list of optimization conditions').grid(column=0,row=0)
            tk.Button(optimization_screen_frame,text='Go',command=lambda: show_entry_fields(make_tray_copy=True)).grid(row=0,column=1,sticky='N,W')

            ttk.Label(optimization_screen_frame,text='Or for complete manual input, look up a reference condition from a screen in CrystalDex').grid(column=0,row=1,sticky='N,W')
            tk.Button(optimization_screen_frame,text='Look up',command=lambda: show_entry_fields(make_tray_copy=False)).grid(row=1,column=1,sticky='N,W')

        def show_entry_fields(make_tray_copy):
            self.clear_widgets()
            self.add_menu()
            self.root.geometry(f'1250x700+{self.screen_width//2-625}+{self.screen_height//2-350}')
            optimization_screen_frame = ttk.Frame(self.root,padding="3 3 12 12")
            optimization_screen_frame.grid(column=0,row=0,sticky='N,W,E,S')
            if not make_tray_copy:
                def choose_ref():
                    for crystal_screen in self.crystal_screens.keys():
                        self.crystal_screen_values.append(crystal_screen)
                        self.crystal_screen_symbols[crystal_screen] = crystal_screen.split('__')[1]
                        crystal_screens_label = ttk.Label(optimization_screen_frame,text='Available screens:')
                        crystal_screens_label.grid(row=1,column=0)
                        crystal_screen_var = tk.StringVar(value=self.crystal_screen_values)
                        crystal_screens_listbox = tk.Listbox(optimization_screen_frame,listvariable=crystal_screen_var,height=5,width=20)
                        crystal_screens_listbox.grid(row=2,column=0)
                    def select_and_continue():
                        lookup_conditions = []
                        for condition in self.crystal_screens.get(crystal_screen):
                            lookup_conditions.append(condition)
                        lookup_conditions_var = tk.StringVar(value=lookup_conditions)
                        self.crystal_screen = crystal_screen_var.get()
                        lookup_listbox = tk.Listbox(optimization_screen_frame,listvariable=lookup_conditions_var,height=20,width=100)
                        lookup_listbox.grid(row=0,column=1,columnspan=4,rowspan=3)

                        def select_condition_to_optimize(event):
                            selection = lookup_listbox.curselection()
                            if selection!=-1:
                                index = selection[0]
                                print(f'index: {index}')
                                selected_index.set(index)
                                self.selected_condition.set(f'{lookup_conditions[index]}')
                                self.clear_widgets()
                                self.add_menu()
                                self.root.geometry(f'1250x700+{self.screen_width//2-625}+{self.screen_height//2-350}')
                                optimization_screen_frame = ttk.Frame(self.root,padding="3 3 12 12")
                                optimization_screen_frame.grid(column=0,row=0,sticky='N,W,E,S')
                                show_entry_fields(make_tray_copy=make_tray_copy)

                        lookup_listbox.bind('<<ListboxSelect>>',select_condition_to_optimize)

                    tk.Button(optimization_screen_frame,text='Select crystal screen and continue',command=select_and_continue).grid(row=3,column=0)

                if self.selected_condition.get() =='':
                    choose_ref()
                else:
                    reference_label = ttk.Label(optimization_screen_frame,text=f'Reference condition: {self.selected_condition.get()}')
                    reference_label.grid(row=0,column=0)

                    tk.Button(optimization_screen_frame,text='Look up new reference',command=lambda: show_entry_fields(make_tray_copy=False)).grid(row=0,column=1,sticky='N,W')

                    ttk.Label(optimization_screen_frame,text="Write a condition and the start, stop, and step concentrations/pH you'd like to iterate that condition over for both the x and y directions. You can populate up to 96 wells.").grid(row=5,column=0,columnspan=5,sticky='N,W')
                    ttk.Label(optimization_screen_frame,text='Please enter in the relevant information for each condition. Ensure that the same number of steps will be generated for your pH and condition settings!').grid(row=6,column=0,columnspan=5,sticky='N,W')

                    ttk.Label(optimization_screen_frame,text='Steps (optional, default is 1):').grid(row=8,column=0)
                    steps_var = tk.StringVar()
                    steps_entry = tk.Entry(optimization_screen_frame,textvariable=steps_var)
                    steps_entry.grid(row=8,column=1)

                    ttk.Label(optimization_screen_frame,text='pH Start (leave empty if not tracked):').grid(row=8,column=1)
                    pH_start_var = tk.StringVar()
                    pH_start_entry = tk.Entry(optimization_screen_frame,textvariable=pH_start_var)
                    pH_start_entry.grid(row=8,column=2)
                    ttk.Label(optimization_screen_frame,text='pH Stop:').grid(row=8,column=3)
                    pH_stop_var = tk.StringVar()
                    pH_stop_entry = tk.Entry(optimization_screen_frame,textvariable=pH_stop_var)
                    pH_stop_entry.grid(row=8,column=4)

                    ttk.Label(optimization_screen_frame,text='Buffer:').grid(row=9,column=0)
                    buffer_var = tk.StringVar()
                    buffer_entry = tk.Entry(optimization_screen_frame,textvariable=buffer_var)
                    buffer_entry.grid(row=9,column=1)
                    ttk.Label(optimization_screen_frame,text='Buffer Concentration Start (Molar):').grid(row=9,column=2)
                    buffer_start_var = tk.StringVar()
                    buffer_start_entry = tk.Entry(optimization_screen_frame,textvariable=buffer_start_var)
                    buffer_start_entry.grid(row=9,column=3)
                    ttk.Label(optimization_screen_frame,text='Buffer Concentration Stop (normally same as Start):').grid(row=9,column=4)
                    buffer_stop_var = tk.StringVar()
                    buffer_stop_entry = tk.Entry(optimization_screen_frame,textvariable=buffer_stop_var)
                    buffer_stop_entry.grid(row=9,column=5)
                    buffer_weight_percent_var = tk.BooleanVar(value=False)
                    buffer_weight_percent_checkbutton = tk.Checkbutton(optimization_screen_frame,text='weight percent',variable=buffer_weight_percent_var,onvalue=True,offvalue=False)
                    buffer_weight_percent_checkbutton.grid(row=9,column=6)

                    ttk.Label(optimization_screen_frame,text='Ingredient 1:').grid(row=10,column=0)
                    ingredient1_var = tk.StringVar()
                    ingredient1_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient1_var)
                    ingredient1_entry.grid(row=10,column=1)
                    ttk.Label(optimization_screen_frame,text='Ingredient Concentration Start (Molar):').grid(row=10,column=2)
                    ingredient1_start_var = tk.StringVar()
                    ingredient1_start_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient1_start_var)
                    ingredient1_start_entry.grid(row=10,column=3)
                    ttk.Label(optimization_screen_frame,text='Ingredient Concentration Stop:').grid(row=10,column=4)
                    ingredient1_stop_var = tk.StringVar()
                    ingredient1_stop_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient1_stop_var)
                    ingredient1_stop_entry.grid(row=10,column=5)

                    ttk.Label(optimization_screen_frame,text='Ingredient 2:').grid(row=11,column=0)
                    ingredient2_var = tk.StringVar()
                    ingredient2_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient2_var)
                    ingredient2_entry.grid(row=11,column=1)
                    ttk.Label(optimization_screen_frame,text='Ingredient Concentration Start (Molar):').grid(row=11,column=2)
                    ingredient2_start_var = tk.StringVar()
                    ingredient2_start_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient2_start_var)
                    ingredient2_start_entry.grid(row=11,column=3)
                    ttk.Label(optimization_screen_frame,text='Ingredient Concentration Stop:').grid(row=11,column=4)
                    ingredient2_stop_var = tk.StringVar()
                    ingredient2_stop_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient2_stop_var)
                    ingredient2_stop_entry.grid(row=11,column=5)

                    ttk.Label(optimization_screen_frame,text='Ingredient 3:').grid(row=12,column=0)
                    ingredient3_var = tk.StringVar()
                    ingredient3_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient3_var)
                    ingredient3_entry.grid(row=12,column=1)
                    ttk.Label(optimization_screen_frame,text='Ingredient Concentration Start (Molar):').grid(row=12,column=2)
                    ingredient3_start_var = tk.StringVar()
                    ingredient3_start_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient3_start_var)
                    ingredient3_start_entry.grid(row=12,column=3)
                    ttk.Label(optimization_screen_frame,text='Ingredient Concentration Stop:').grid(row=12,column=4)
                    ingredient3_stop_var = tk.StringVar()
                    ingredient3_stop_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient3_stop_var)
                    ingredient3_stop_entry.grid(row=12,column=5)
                    
                    conditions_var = tk.StringVar(value=conditions)
                    self.selected_condition = tk.StringVar()
                    conditions_listbox = tk.Listbox(optimization_screen_frame,listvariable=conditions_var,height=25,width=150)
                    conditions_listbox.grid(row=13,column=0,columnspan=3)

                    new_condition_instructions = [[ingredient1_var.get(),ingredient1_start_var.get(),ingredient1_stop_var.get(),ingredient1_weight_percent_var.get()],[ingredient2_var.get(),ingredient2_start_var.get(),ingredient2_stop_var.get(),ingredient2_weight_percent_var.get()],[ingredient3_var.get(),ingredient3_start_var.get(),ingredient3_stop_var.get(),ingredient3_weight_percent_var.get()],[buffer_var.get(),buffer_start_var.get(),buffer_stop_var.get(),buffer_weight_percent_var.get()]]
                    new_conditions = ['' for _ in range(96)]

                    def save_condition_settings():
                        steps = 1
                        pH_start = -99
                        pH_stop = -99
                        pH_step = 0
                        if pH_start_var!=-99:
                            pH_start = float(pH_start_var.get())
                            if pH_stop_var!=-99:
                                pH_stop = float(pH_stop_var.get())
                                if steps_var:
                                    steps = int(steps_var.get())
                                    if pH_stop>pH_start:
                                        pH_step = round((pH_stop-pH_start)/(steps-1),2)
                            else:
                                messagebox.showerror(title='No pH Stop',message=f"Please enter a pH Stop. This is the pH you'd like your current set of conditions to end at.")

                        for new_condition_number in range(steps):
                            for new_ingredient in new_condition_instructions:
                                if '' not in [new_ingredient[0],new_ingredient[1],new_ingredient[2]]:
                                    new_ingredient_id = new_ingredient[0]
                                    new_condition_start = float(new_ingredient[1])
                                    new_condition_stop = float(new_ingredient[2])
                                    new_condition_step = round((new_condition_stop-new_condition_start)/(steps-1),2)
                                    new_condition_concentration = new_condition_number*new_condition_step+new_condition_start
                                    if not new_ingredient[3]:
                                        new_conditions[new_condition_number] += f'{new_condition_concentration} M {new_ingredient_id}, '
                                    else:
                                        new_conditions[new_condition_number] += f'{new_condition_concentration} % {new_ingredient_id}'
                            if pH_start!=-99:
                                new_conditions[new_condition_number] += f'pH {new_condition_number*pH_step+pH_start}'

                        for condition in range(len(conditions)):
                            if len(new_conditions)>0:
                                if conditions[condition] == '':
                                    conditions[condition] = f'{condition+1}. {new_conditions.pop(0)}'
                                    conditions_listbox.delete(condition)
                                    conditions_listbox.insert(condition, conditions[condition])
                    
                    tk.Button(optimization_screen_frame,text='Add selection to optimization',command=save_condition_settings).grid(row=50,column=0)

                    tk.Button(optimization_screen_frame,text='Finish optimization screen',command=save).grid(row=51,column=0)

            elif make_tray_copy:
                w = 8
                link = 'https://hamptonresearch.com/make-tray.php'
                webbrowser.get('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe %s').open(link)
                time.sleep(1)
                self.refocus()
                self.root.geometry(f'{self.screen_width//2}x{self.screen_height}+0+0')#Not sure why this line doesn't work
                optimization_screen_frame = ttk.Frame(self.root,padding="3 3 12 12")
                optimization_screen_frame.grid(column=0,row=0,sticky='N,W,E,S')#the full sticky means this fills the master frame.
                ttk.Label(optimization_screen_frame,text="Copy the conditions seen in Make Tray as well as possible.").grid(row=5,column=0,columnspan=5,sticky='N,W')

                ttk.Label(optimization_screen_frame,text='Lowest pH (leave empty if not tracked):').grid(row=8,column=1)
                pH_start_var = tk.StringVar()
                pH_start_entry = tk.Entry(optimization_screen_frame,textvariable=pH_start_var,width=w)
                pH_start_entry.grid(row=8,column=2)
                ttk.Label(optimization_screen_frame,text='Highest pH:').grid(row=8,column=3)
                pH_stop_var = tk.StringVar()
                pH_stop_entry = tk.Entry(optimization_screen_frame,textvariable=pH_stop_var,width=w)
                pH_stop_entry.grid(row=8,column=4)

                ttk.Label(optimization_screen_frame,text='Buffer:').grid(row=9,column=0)
                buffer_var = tk.StringVar()
                buffer_entry = tk.Entry(optimization_screen_frame,textvariable=buffer_var)
                buffer_entry.grid(row=9,column=1)
                ttk.Label(optimization_screen_frame,text='Lowest Buffer Concentration:').grid(row=9,column=2)
                buffer_start_var = tk.StringVar()
                buffer_start_entry = tk.Entry(optimization_screen_frame,textvariable=buffer_start_var,width=w)
                buffer_start_entry.grid(row=9,column=3)
                ttk.Label(optimization_screen_frame,text='Highest Buffer Concentration (normally same as lowest):').grid(row=9,column=4)
                buffer_stop_var = tk.StringVar()
                buffer_stop_entry = tk.Entry(optimization_screen_frame,textvariable=buffer_stop_var,width=w)
                buffer_stop_entry.grid(row=9,column=5)
                buffer_weight_percent_var = tk.BooleanVar(value=False)
                buffer_weight_percent_checkbutton = ttk.Checkbutton(optimization_screen_frame,text='weight percent',variable=buffer_weight_percent_var,onvalue=True,offvalue=False)
                buffer_weight_percent_checkbutton.grid(row=9,column=6)

                ttk.Label(optimization_screen_frame,text='Ingredient 1:').grid(row=10,column=0)
                ingredient1_var = tk.StringVar()
                ingredient1_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient1_var)
                ingredient1_entry.grid(row=10,column=1)
                ttk.Label(optimization_screen_frame,text='Ingredient Lowest Concentration:').grid(row=10,column=2)
                ingredient1_start_var = tk.StringVar()
                ingredient1_start_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient1_start_var,width=w)
                ingredient1_start_entry.grid(row=10,column=3)
                ttk.Label(optimization_screen_frame,text='Ingredient Highest Concentration:').grid(row=10,column=4)
                ingredient1_stop_var = tk.StringVar()
                ingredient1_stop_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient1_stop_var,width=w)
                ingredient1_stop_entry.grid(row=10,column=5)
                ingredient1_weight_percent_var = tk.BooleanVar(value=False)
                ingredient1_weight_percent_checkbutton = ttk.Checkbutton(optimization_screen_frame,text='weight percent',variable=ingredient1_weight_percent_var,offvalue=False,onvalue=True)
                ingredient1_weight_percent_checkbutton.grid(row=10,column=7)

                ttk.Label(optimization_screen_frame,text='Ingredient 2:').grid(row=11,column=0)
                ingredient2_var = tk.StringVar()
                ingredient2_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient2_var)
                ingredient2_entry.grid(row=11,column=1)
                ttk.Label(optimization_screen_frame,text='Ingredient Lowest Concentration:').grid(row=11,column=2)
                ingredient2_start_var = tk.StringVar()
                ingredient2_start_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient2_start_var,width=w)
                ingredient2_start_entry.grid(row=11,column=3)
                ttk.Label(optimization_screen_frame,text='Ingredient Highest Concentration:').grid(row=11,column=4)
                ingredient2_stop_var = tk.StringVar()
                ingredient2_stop_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient2_stop_var,width=w)
                ingredient2_stop_entry.grid(row=11,column=5)
                ingredient2_weight_percent_var = tk.BooleanVar(value=False)
                ingredient2_weight_percent_checkbutton = ttk.Checkbutton(optimization_screen_frame,text='weight percent',variable=ingredient2_weight_percent_var,offvalue=False,onvalue=True)
                ingredient2_weight_percent_checkbutton.grid(row=11,column=7)

                ttk.Label(optimization_screen_frame,text='Ingredient 3:').grid(row=12,column=0)
                ingredient3_var = tk.StringVar()
                ingredient3_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient3_var)
                ingredient3_entry.grid(row=12,column=1)
                ttk.Label(optimization_screen_frame,text='Ingredient Lowest Concentration:').grid(row=12,column=2)
                ingredient3_start_var = tk.StringVar()
                ingredient3_start_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient3_start_var,width=w)
                ingredient3_start_entry.grid(row=12,column=3)
                ttk.Label(optimization_screen_frame,text='Ingredient Highest Concentration:').grid(row=12,column=4)
                ingredient3_stop_var = tk.StringVar()
                ingredient3_stop_entry = tk.Entry(optimization_screen_frame,textvariable=ingredient3_stop_var,width=w)
                ingredient3_stop_entry.grid(row=12,column=5)
                ingredient3_weight_percent_var = tk.BooleanVar(value=False)
                ingredient3_weight_percent_checkbutton = ttk.Checkbutton(optimization_screen_frame,text='weight percent',variable=ingredient3_weight_percent_var,offvalue=False,onvalue=True)
                ingredient3_weight_percent_checkbutton.grid(row=12,column=7)
                
                ttk.Label(optimization_screen_frame,text='Virtual Crystal Screen').grid(row=14,column=2,sticky='W,E')
                vcs = tk.Frame(optimization_screen_frame,width=600,height=400,relief='raised',background='white',border=4,highlightcolor='white')
                vcs.grid(row=15,column=1,columnspan=6)
                quad1 = tk.Frame(vcs,width=300,height=200,borderwidth=2,background='blue')
                quad1.grid(row=0,column=0)
                quad2 = tk.Frame(vcs,width=300,height=200,borderwidth=2)
                quad2.grid(row=1,column=0)
                quad3 = tk.Frame(vcs,width=300,height=200,borderwidth=2)
                quad3.grid(row=0,column=1)
                quad4 = tk.Frame(vcs,width=300,height=200,borderwidth=2)
                quad4.grid(row=1,column=1)

                new_conditions = ['' for _ in range(96)]
                self.condition = -1

                def save_condition_settings():                        
                    if self.quad == 1:
                        self.quad = 2
                        quad1.configure(background='gray')
                        quad2.configure(background='blue')
                    elif self.quad ==2:
                        self.quad = 3
                        quad2.configure(background='gray')
                        quad3.configure(background='blue')
                    elif self.quad ==3:
                        self.quad = 4
                        quad3.configure(background='gray')
                        quad4.configure(background='blue')
                    elif self.quad == 4:
                        self.quad = 5

                    steps = 6
                    pH_start = -99
                    pH_stop = -99
                    pH_step = 0
                    pH_steps = 4
                    if pH_start_var.get() !='':
                        pH_start = float(pH_start_var.get())
                        if pH_stop_var.get()!='':
                            pH_stop = float(pH_stop_var.get())
                            if pH_stop>pH_start:
                                pH_step = round((pH_stop-pH_start)/(pH_steps-1),2)
                        else:
                            messagebox.showerror(title='No pH Stop',message=f"Please enter a highest pH.")

                    new_condition_instructions = [[ingredient1_var.get(),ingredient1_start_var.get(),ingredient1_stop_var.get(),ingredient1_weight_percent_var.get()],[ingredient2_var.get(),ingredient2_start_var.get(),ingredient2_stop_var.get(),ingredient2_weight_percent_var.get()],[ingredient3_var.get(),ingredient3_start_var.get(),ingredient3_stop_var.get(),ingredient3_weight_percent_var.get()],[buffer_var.get(),buffer_start_var.get(),buffer_stop_var.get(),buffer_weight_percent_var.get()]]
                    for new_condition_number in range(steps):
                        for new_pH_number in range(pH_steps):
                            self.condition = self.condition+1
                            for new_ingredient in new_condition_instructions:
                                if '' not in [new_ingredient[0],new_ingredient[1],new_ingredient[2]]:
                                    new_ingredient_id = new_ingredient[0]
                                    new_condition_start = float(new_ingredient[1])
                                    new_condition_stop = float(new_ingredient[2])
                                    new_condition_step = round((new_condition_stop-new_condition_start)/(steps-1),2)
                                    new_condition_concentration = new_condition_number*new_condition_step+new_condition_start
                                    if not new_ingredient[3]:#if not measured by weight percent:
                                        new_conditions[self.condition] += f'{new_condition_concentration} M {new_ingredient_id}, '
                                    else:
                                        new_conditions[self.condition] += f'{new_condition_concentration} % {new_ingredient_id}, '
                            if pH_start != -99:
                                new_conditions[self.condition] += f'pH {new_pH_number*pH_step+pH_start}'
                            
                    if self.quad == 5:
                        quads = [[],[],[],[]]
                        condition = -1
                        for quad in range(len(quads)):
                            for _ in range(len(conditions)//4):
                                condition += 1
                                quads[quad].append(new_conditions[condition])
                                
                        for step in range(steps):
                            for quad in range(4):
                                for pH_step in range(4):
                                    if quad == 0:
                                        conditions[step*8+pH_step] = quads[quad][step*4+pH_step]
                                    elif quad ==1:
                                        conditions[step*8+pH_step+4] = quads[quad][step*4+pH_step]
                                    elif quad ==2:
                                        conditions[step*8+pH_step+48] = quads[quad][step*4+pH_step]
                                    elif quad ==3:
                                        conditions[step*8+pH_step+52] = quads[quad][step*4+pH_step]
                        review_make_tray_copy()

                save_condition_settings_button = tk.Button(optimization_screen_frame,text='Add selection \nto optimization',command=save_condition_settings)
                save_condition_settings_button.grid(row=13,column=0)

                def review_make_tray_copy():
                    self.clear_widgets()
                    self.add_menu()
                    self.root.geometry(f'1250x700+{self.screen_width//2-625}+{self.screen_height//2-350}')
                    optimization_screen_frame = ttk.Frame(self.root,padding="3 3 12 12")
                    optimization_screen_frame.grid(column=0,row=0,sticky='N,W,E,S')

                    listbox_label = ttk.Label(optimization_screen_frame,text='Review and correct generated conditions:')
                    listbox_label.grid(row=2,column=0)
                    listbox_values = [f"[{condition+1}]: {conditions[condition]}" for condition in range(len(conditions))]
                    condition_var = tk.StringVar(value=listbox_values)
                    conditions_listbox = tk.Listbox(optimization_screen_frame,listvariable=condition_var,height=25,width=150)
                    conditions_listbox.grid(row=3,column=0,columnspan=3)
                    
                    edited_condition = tk.StringVar()
                    condition_entry = tk.Entry(optimization_screen_frame, textvariable=edited_condition, width=150)
                    condition_entry.grid(row=4, column=0, columnspan=3)

                    selected_index = tk.IntVar(value=-1)

                    def select_condition(event):
                        selection = conditions_listbox.curselection()
                        if selection:
                            index = selection[0]
                            selected_index.set(index)
                            edited_condition.set(f'{conditions[index]}')

                    conditions_listbox.bind('<<ListboxSelect>>',select_condition)

                    def overwrite():
                        index = selected_index.get()
                        #print(f'index: {index}')
                        if index >= 0:
                            text = edited_condition.get()
                            condition = index
                            conditions[condition] = text
                            conditions_listbox.delete(index)
                            conditions_listbox.insert(index, f'[{condition+1}]: {text}')
                            selected_index.set(index+1)
                            if index !=95:
                                edited_condition.set(f'{conditions[index+1]}')
                            else:
                                edited_condition.set(f'{conditions[0]}')
                            
                    tk.Button(optimization_screen_frame,text='overwrite',command=overwrite).grid(row=4,column=3)

                    tk.Button(optimization_screen_frame,text='Finish optimization screen',command=save).grid(row=5,column=3)

        def save():
            self.crystal_screens[f'{self.long_name}__{self.two_code}'] = conditions
            #print(f'crystal_screens: {self.crystal_screens}')
            with open(crystal_screens_path, "w") as c:
                json.dump(self.crystal_screens, c)
            self.splash()
            threading.Thread(target=self.Box_Save,daemon=True).start()
            self.startup()
