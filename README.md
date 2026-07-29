# Bizhub_Printer_Counter
This C# program reads data from the Bizhub web interface and writes the counter values to an XLSX file based on the printers' serial numbers (SN).

Requirements
There is a text file named IPs.txt. Add the IP addresses of your printers to this file.
The program will create an "Archiv" directory. There you need to upload the sample xlsx named "Sablon".

The file should be formatted like this:
XXX.XXX.XXX.XXX
XXX.XXX.XXX.XXX

Compatible Bizhub models
C364e
C308
308e
C300l
C250i

It is probably compatible with additional Bizhub models, but these are the ones that have been tested and confirmed to work.
