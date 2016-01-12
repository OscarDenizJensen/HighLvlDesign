import xlsxwriter
from jnpr.junos import Device

##############################
##      Connect to Device   ##
##############################

#Enter Device Credentials
dev = Device(host='172.16.18.1', user='oscar', password='Password')

#Open Connection
dev.open()

##############################
##  Get Config from Device  ##
##############################
def get_conf():
    #Get Config in SET format
    return dev.cli("show configuration | display set")

config=str(get_conf())

#Enter File name
file_name=raw_input("Enter File Name including extention: ")

#Write Config to Text File
with open(file_name, "w") as text_file:
        text_file.write(config)
        text_file.close()


##############################################
## Create a workbook and add a worksheet.   ##
##############################################
workbook = xlsxwriter.Workbook('Test-Document.xlsx')
worksheet = workbook.add_worksheet()

#Open Document in read forat
config=open(file_name,"r")

#Transfer text to a file
config_list=[]
for x in config:
    if "set" in x:
        print x
        config_list.append(x)

#Create new list to put updated list. Updated list doesn't have repetition
updated_list=[]

##################################################
##  Get rid of Duplicate info in the First List ##
##################################################
record=dict()
for line in config_list:
    #Split lines
    temp = line.split()
    #Check for duplicates
    if temp[0] in record:
        key = temp[0]
        temp[0] = " "           #Replace duplicate with given value
        for i in range(1,len(temp)):
            if temp[i] == record[key][i-1]:
                temp[i] = " "   #Replace duplicate with given value
            else:
                break           #Stop at first different info

    updated_list.append(temp)    #Write new info Updated list

    temp = line.split()
    record[temp[0]] = temp[1:]

#############################
#   WRITING TO EXCEL        #
#############################

#Set start values for row and column
row=0
col=-1  #Column is set -1 so it won't enter 'Set' command into excel

#Transfer config into Excel
for line in updated_list:
    for statement in line:
        #Final statement in a line
        if statement==line[(len(line)-1)]:
            worksheet.write(row,col,statement)
            row+=1  #Skip to next row if it is final statement
            col=-1

        #New Hierarchy
        elif statement== line[1] and statement != " ":
            row+=1              #New row to seperate from previous hierarchy
            worksheet.write(row,col,statement)
            col+=1

        #Anything Else
        else:
            worksheet.write(row,col,statement)
            col+=1  #Move to next column if it is not final statement

#Close workbook
workbook.close()