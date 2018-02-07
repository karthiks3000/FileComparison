# This program compares 2 folders and spits out an excel with the results

import Tkinter,tkFileDialog
import os
import hashlib
import sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font,colors

book = Workbook()

font = Font(name='Calibri',
size=11,
bold=True,
italic=False,
vertAlign=None,
underline='none',
strike=False,
color=colors.RED)

# function to calculate the checksum value of the provided file path
def md5Checksum(filePath):
    with open(filePath, 'rb') as fh:
        m = hashlib.md5()
        while True:
            data = fh.read(8192)
            if not data:
                break
            m.update(data)
        return m.hexdigest()

# code starts here

root = Tkinter.Tk()
root.withdraw()
folder_size_source = 0
folder_size_destination = 0


# Selects the Source Folder
sourceFolder = tkFileDialog.askdirectory()
# sourceFolder = 'C:\Users\subramak\Downloads\Test'

print('Source Path: ' + sourceFolder)
print('')

# Selects the Destination Folder
destinationFolder = tkFileDialog.askdirectory()
# destinationFolder = 'C:\Users\subramak\Downloads\Test - Copy'

print('Destination Path: ' + destinationFolder)
print('')
   

sheet1 = book.create_sheet('Results')

sheet1.cell(1,1,'Source Folder: ')

sheet1.cell(1,2, sourceFolder)

sheet1.cell(2,1,'Destination Folder:')
sheet1.cell(2,2, destinationFolder)

sheet1.cell(5,1,'File Comparison Results:')

sheet1.cell(7,1,'Source File Path')
sheet1.cell(7,2,'Source File MD5')
sheet1.cell(7,3,'Destination File Path')
sheet1.cell(7,4,'Destination File MD5')
sheet1.cell(7,5,'Result')

i = 8

# for each file in source, compare with destination   
for (path, dirs, files) in os.walk(sourceFolder):
    for file in files:
        source_filename = os.path.join(path, file)          
        sheet1.cell(i,1,source_filename)
        
        destination_filename = os.path.join(destinationFolder, os.path.relpath(path,sourceFolder))
        destination_filename = os.path.join(destination_filename,file)
        
        folder_size_source += os.path.getsize(source_filename)
        
        if(os.path.isfile(destination_filename)):    

            sheet1.cell(i,3,destination_filename)   

            folder_size_destination += os.path.getsize(destination_filename)    

            SourceCheckSum = md5Checksum(source_filename)
            sheet1.cell(i,2,SourceCheckSum)   
            
            DestinationCheckSum = md5Checksum(destination_filename)
            sheet1.cell(i,4,DestinationCheckSum)   
            
            #Compare the checksum of Source and Files
            if(SourceCheckSum != DestinationCheckSum):
                sheet1.cell(i,5,'Checksum does not match')
            else:
                sheet1.cell(i,5,'Checksum matched')
                      
        else:
            sheet1.cell(i,5,destination_filename + ' does not exist')            
        
        i = i+1
            
i=i+3
sheet1.cell(i,1,'Size of files found in source')
sheet1.cell(i,2,'%0.1f MB' % (folder_size_source/(1024*1024.0)))

i=i+1
sheet1.cell(i,1,'Size of files found in destination')
sheet1.cell(i,2,'%0.1f MB' % (folder_size_destination/(1024*1024.0)))        
   
print('')
print("Source Folder Size: %0.1f MB" % (folder_size_source/(1024*1024.0)))
print("Destination Folder Size = %0.1f MB" % (folder_size_destination/(1024*1024.0)))

book.save('Result.xlsx')

