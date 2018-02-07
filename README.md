# File Comparison
This simple script allows you to specify a source and a destination folder and it compares all files recursively within them.
The result is an excel file with the details of the comparison.

Each file from the source is checked to see if a corresponding file with the same name exists in the destination.
If found, both files are compared by generating an MD5 for each of them.

Dependencies: hashlib,openpyxl