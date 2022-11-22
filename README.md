# Active scripts for Excel
A collection of scripts utilizing Excel as UI and calling Python scripts from within VBA with aid of Xlwings.

## Container Booking Sheet with Excel UI
Main program is the Container Bookings Sheet that based on information recorded the INFO-sheet can create the following files:

- *Estimated Load List (ELL), later used to create COPARN/COPRAR/BAPLIE files.*
- *Container Booking Forecast (CBF)*
- *Dangerous Cargo Manifest (DCM)*
- *Reefer Log Book (RLB)*
- *Pre export file to Main Line Operators (MLO)*

Container Booking Sheets supports data validation.  
Example file runs when script is called directly, it will save files to the desktop. Make sure that Excel is configured as below.

## PDF-parser for FITOR with Excel UI
Similar to Container Booking Sheet this PDF-parser uses Excel as UI and Xlwings to call Python and work with the Excel file.
Program parses a Cargo Manifest in PDF-format from FITOR for information.

Runtime is 10 ~ 30 seconds based on size of the PDF-file as there are a multitude of for loops running to collect unorganized data over several pages.
Example file runs when script is called directly, it will save file to the desktop. Make sure Excel is configured as below.

## Excel set up
In order to set up `Xlwings` as an add-in in Excel you need to do the following:

Click on `Developer`-ribbon --> `Visual Basic` or press Alt + F11 to enter the developer pane.
From the menu choose `Tools` --> `References` and then tick `Xlwings`.

![Image](/images/developer_addin_xlwings.png)

You also need to configure path to Python, either in Excel as below,

![Image](/images/xlwings_configuration.png)

or in .xlwings config file in your user path:  
`.xlwings\xlwings.conf`

## External packages
`Xlwings` `Pandas` `Numpy` `PymuPDF`

requirements.txt includes linting package `Black` but not used in production.

## TODO
- [ ] Split `fitor_pdf_parser` into several files and add documentation.
## License
MIT
