import xlrd
import sys

if(len(sys.argv) != 1):
    
    f = open("contacts.vcf", "w")

    location = (sys.argv[1])

    workbook = xlrd.open_workbook(location)
    sheet = workbook.sheet_by_index(0)

    rows = sheet.nrows
    cols = sheet.ncols


    for i in range(rows):
        to_write = ''
        if(i!=0):
            to_write = '\n'
        to_write = to_write+'BEGIN:VCARD\nVERSION:3.0\nN:_CODE_;'+str(sheet.cell_value(i,0))+';;;\nFN:_CODE_+'+str(sheet.cell_value(i,0))+'\nTEL;TYPE=CELL:'+str(int(sheet.cell_value(i,1)))+'\nTEL;TYPE=CELL:'+str(int(sheet.cell_value(i,1)))+'\nEND:VCARD'
        f.write(to_write)


    f.close()


else:

    print("Please Type the name of File as Argument")