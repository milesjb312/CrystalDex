CrystalDex is an application for making manual crystal tray checking easier to track by creating virtual crystal trays. It can upload into Microsoft Box and can access the SeBaView software to obtain images.

Note that this software currently works only with the screensize of the Moody Lab microscope computer.

Functional goals of this program:
    1) To make it easy to put both pictures and descriptions of every crystal we find immediately into the local server or into Box without having to open up Box and copy things to the Excel sheet that normally requires very repetitive data entry (and in which we often miss things).
        - This has been accomplished by using the Box SDK with Python to access Box and by creating a Tkinter GUI that is easy to interact with. The GUI contains reference fields to be filled out and buttons for operating the microscope in an integrated fashion.
    2) To incorporate the information from all previous experimental steps into a single place so that we can track easily (and with less tedium) how our experiments are going.
        - This has been accomplished by creating a background database that includes all of the crystal conditions that we normally use, as well as GUI-led steps for uploading new crystal conditions and crystal optimization conditions (which are often difficult to keep track of on paper). The GUI leads the user to operate a web-sourced crystal optimization condition generator (https://hamptonresearch.com/make-tray.php) and then prompts the user to enter the generated conditions by hand into the database to be referenced later on in the Excel-sheet editing steps.
        https://docs.python.org/3/library/tkinter.html

Currently, CrystalDex does the following:
    1) Indexes Trays by:
        - Requesting information about the protein construct, crystallization screen, and date, and creating a new Excel worksheet within a workbook that contains all virtual crystal trays previously indexed.
        - Automatically opening the SeBaView software to allow the user to see the camera image.
        - Requesting information about the well and subwell currently focused.
        - Allowing the user to size the crystal using a click-and-drag sizing tool.
        - Automatically saving the image with a conserved naming system and putting the information into a virtual tray along with a link to the picture.
        - Uploading the entire workbook and all images into the server and (optionally) into Box.
    2) Generates a Crystal Sendoff sheet during the Harvesting step by:
        - Referencing the internal database containing all conditions, either commercial or optimized and automatically filling out the Crystal Sendoff sheet with all pertinent data, including crystal size and any notes.

If you're wondering about the purposes behind the existence of certain folders/files, here's the reason for each:
- Crystal_Pictures folder: holds all the crystal pictures and is searched by CrystalDex each time it tries to upload to Box. It is/should be fully mirrored into Box each time CrystalDex tries to upload to Box.0
- Crystal_Screens.json: holds all the information about which conditions apply to which crystal screens/optimization screens and is searched by CrystalDex whenever a Crystal_Sendoff_Sheet is updated.
- Crystal_Sendoff_Sheet.xlsx: holds information about the current set of crystals to be sent to an X-Ray source. It is updated when users are harvesting crystals and searched whenever CrystalDex uploads to Box.
- crystaldex_icon.ico: the icon for the app.
- crystaldex_icon.png: the icon that shows up on the corner of the screen in the GUI
- CrystalDex_Library_Mastercopy.xlsx: a backup that allows CrystalDex to rewrite the basic format of the CrystalDex library in case it is destroyed or corrupted.
- CrystalDex_Library.xlsx: an Excel workbook that contains separate worksheets, where each worksheet is a virtual crystal screen with information about the conditions, the date set, and all the crystals imaged so far. It contains links to each image for convenience.
- CrystalDex_splash.png: the splash image that is shown when CrystalDex is booting up.
- SeBaView_path_file.json: a file that says where SeBaView is stored on your computer. If it fails, the user is prompted to select the SeBaView application, and the file is overwritten.
- CrystalDex.gitignore: this file instructs GitHub to ignore the Box user authentication files generated whenever you first use CrystalDex or whenever Box fails to connect.
- CrystalDex.py: this is where all the code for CrystalDex is stored
- main.py: this is an entry point for the CrystalDex application, used when PyInstaller is trying to create an executable. If you're running the app from a terminal, this is the script you want to run.
- main.spec: this file contains specifications for PyInstaller, such as which libraries to download to make a CrystalDex executable operational.
- requirements.txt: this file isn't used for anything as far as I'm aware, I just generated it using pip > freeze to show what libraries/versions I need to make CrystalDex work.

Future updates will include a queue so that you can still attempt to update Box files even when someone else is currently editing them (currently, the upload fails and raises an error).