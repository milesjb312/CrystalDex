CrystalDex is an application for making manual crystal tray checking easier to track by creating virtual crystal trays. It uploads into Microsoft Box and can access the SeBaView software to obtain images.

Note that this software currently works only with the screensize of the Moody Lab microscope computer.

Functional goals of this program:
    1) To make it easy to put both pictures and descriptions of every crystal we find immediately into Box without having to open up Box and copy things to the Excel sheet that normally requires very repetitive data entry (and in which we often miss things).
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
        - Uploading the entire workbook and all images into Box.
    2) Generates a Crystal Sendoff sheet during the Harvesting step by:
        - Referencing the internal database containing all conditions, either commercial or optimized and automatically filling out the Crystal Sendoff sheet with all pertinent data, including crystal size and any notes.