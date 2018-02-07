# This program compares 2 folders and spits out an excel with the results

import Tkinter,tkFileDialog
import filecmp
import os
import hashlib
import optparse
import os.path
import sys
import xlwt

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
print('Source Path: ' + sourceFolder)
print('')

# Selects the Destination Folder
destinationFolder = tkFileDialog.askdirectory()
print('Destination Path: ' + destinationFolder)
print('')
   
# open a workbook to save the results   
book = xlwt.Workbook(encoding="utf-8")

sheet1 = book.add_sheet("Results")
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')
sheet1.col(0).width = 256 * 70
sheet1.col(1).width = 256 * 50
sheet1.col(2).width = 256 * 70
sheet1.col(3).width = 256 * 50
sheet1.col(4).width = 256 * 70

sheet1.write(0,0,'Source Folder: ',style0)

sheet1.write(0,1, sourceFolder)

sheet1.write(1,0,'Destination Folder:',style0)
sheet1.write(1,1, destinationFolder)

sheet1.write(5,0,'File Comparison Results:',style0)

sheet1.write(7,0,'Source File Path',style0)
sheet1.write(7,1,'Source File MD5',style0)
sheet1.write(7,2,'Destination File Path',style0)
sheet1.write(7,3,'Destination File MD5',style0)
sheet1.write(7,4,'Result',style0)

i = 8

# for each file in source, compare with destination   
for (path, dirs, files) in os.walk(sourceFolder):
    for file in files:
        source_filename = os.path.join(path, file)          
        sheet1.write(i,0,source_filename)
        
        destination_filename = os.path.join(destinationFolder, os.path.relpath(path,sourceFolder))
        destination_filename = os.path.join(destination_filename,file)
        
        folder_size_source += os.path.getsize(source_filename)
        
        if(os.path.isfile(destination_filename)):    

            sheet1.write(i,2,destination_filename)   

            folder_size_destination += os.path.getsize(destination_filename)    

            SourceCheckSum = md5Checksum(source_filename)
            sheet1.write(i,1,SourceCheckSum)   
            
            DestinationCheckSum = md5Checksum(destination_filename)
            sheet1.write(i,3,DestinationCheckSum)   
            
            #Compare the checksum of Source and Files
            if(SourceCheckSum != DestinationCheckSum):
                sheet1.write(i,4,'Checksum does not match')
            else:
                sheet1.write(i,4,'Checksum matched')
                      
        else:
            sheet1.write(i,4,destination_filename + ' does not exist')            
        
        i = i+1
            
i=i+3
sheet1.write(i,0,'Source Folder Size',style0)
sheet1.write(i,1,'%0.1f MB' % (folder_size_source/(1024*1024.0)))

i=i+1
sheet1.write(i,0,'Destination Folder Size',style0)
sheet1.write(i,1,'%0.1f MB' % (folder_size_destination/(1024*1024.0)))        
   
print('')
print("Source Folder Size: %0.1f MB" % (folder_size_source/(1024*1024.0)))
print("Destination Folder Size = %0.1f MB" % (folder_size_destination/(1024*1024.0)))

book.save('Result.xls')

