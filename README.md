# Li6800ExcelExtract
For extracting columns from LI6800 excel files and combining into one complete excel file with a grouping variable. Designed for the fitacis function in the plantecophys package in R

This code is designed to take multiple excel files from the Licor LI6800 and combine them into one excel file with the corresponding file titles. This works best if the title of each file is the same as the grouping variable.

For example, my file for red oak m3 is set as 2025-05-19-2248_rom3_resp.xlsx 
For red oak g2 2025-05-19-2253_rog2_resp.xlsx
This code would combine them with a corresponding column that includes the title of the file.

From there you can use excel to seperate the column by the _ deliniator. This sets up the file for use in the fitacis function of plantecophys R package:
https://www.rdocumentation.org/packages/plantecophys/versions/1.4-6/topics/fitacis

Some important notes: 
	Have all excel (.xlsx) files in a separate folder from the other Licor files. This code is only designed to read .xlsx files.
	Make sure all files have unique names. As the LI6800 uses unique timestamps for each file, it is unlikely that this will be an issue.
	Make sure that the columns you want to extract are present in each file. If they are not, it will fail to run. Delete any empty files. Also ensure each column header is properly typed in the Python code.
	The LI6800 excel files have each column header as three rows. To request a column for export in the code, make sure you put a space between each part of the rows. Example: GasEx A µmol m⁻² s⁻¹     . Not GasExAµmol m⁻² s⁻¹
	
