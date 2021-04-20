# Scannen
This application drives a Canon DR_M260 in order to scan, process(Invoicing &amp; Emailing Delivery Confirmation ) and archive delivery documents.
***********************************************************************************************
Description: This python program would authentify a PDF file using its name. The File name should be formated as scan_"Order#"--"Location". 
pdf. Example: scan_43160--Gro.pdf. After a successful authentification, the program would get a serial number & email address from the 
server the server. It would then Insert the packing slip into the company database. The final database work would be to release Packing slips
into invoices. The program also emails a comfirmation email to the customer. It also has a method that deletes files older than 3 years.
Processed files are first stored locally before being moved to the external Z drive. A window notification appears when the program starts.
In case of an incorrect file name, a window notication would appear. The program also opens the folder where the file is located.
