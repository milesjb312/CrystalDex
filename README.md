CrystalDex is an application for making manual crystal tray checking easier to track by creating virtual crystal trays. It incorporates into Microsoft Box and can access the SeBaView software to obtain images.

Note that this software currently works only with the screensize of the Moody Lab microscope computer.

Functional goals of this program:
    1) To make it easy to put both pictures and descriptions of every crystal we find immediately into Box without having to open up Box and copy things to the Excel sheet that noramlly requires very repetitive data entry (and in which we often miss things).
        - This will be accomplished by using the Box SDK with Python to access Box and by creating a Tkinter GUI that is easy to interact with. The GUI will contain reference fields to 
        be filled out and buttons for operating the microscope in an integrated fashion. It may also contain other operations as described in Goal 2.
    2) To incorporate the information from all the previous experimental steps into a single place so that we can track easily (and with less tedium) how our experiments are going.
        - This will be accomplished by creating a background database that includes all of the crystal conditions that we normally use, as well as GUI-led steps for uploading crystal
        optimization conditions (which are often difficult to keep track of on paper). If possible and deemed necessary, this GUI may lead the user to operate a web-sourced crystal 
        optimization condition generator (https://hamptonresearch.com/make-tray.php) and then will immediately scrape the data into the database to be referenced later on in the 
        Excel-sheet editing steps.
        https://docs.python.org/3/library/tkinter.html

Currently, CrystalDex does the following:
    1) Indexes Trays by:
        - Requesting information about the protein construct, crystallization screen, and date, and creating a new Excel worksheet within a workbook that contains all virtual crystal trays previously indexed.
        - Automatically opening the SeBaView software to allow the user to see the camera image.
        - Requesting information about the well and sub well currently focused.
        - Automatically saving the image with a conserved naming system.
        - Updating the entire workbook into Box.

In the near future, CrystalDex will:
    1) Improve indexing trays by:
        - Allowing users to click-and-drag a sizing tool to measure crystals in a repeatable and consistent manner.
        - Uploading each picture taken directly into Box
        - Creating and posting a link from Box to each picture inside of the appropriate Excel virtual tray.