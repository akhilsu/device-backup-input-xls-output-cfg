from netmiko import ConnectHandler
from netmiko import NetMikoAuthenticationException, NetMikoTimeoutException
import xlrd
from datetime import datetime

print("PROCESS INITIATED..")

openfile = xlrd.open_workbook(r"device_list.xls")
sheet = openfile.sheet_by_name("Sheet1")

print("DATA COPIED SUCCESSFULLY")

for i in range (1,sheet.nrows):
    device = {
        "device_type" : sheet.row(i)[5].value,
        "ip" : sheet.row(i)[2].value,
        "username" : sheet.row(i)[3].value,
        "password" : sheet.row(i)[4].value,
        "port" : 22
    }
    print("DEVICE "+ str(int(sheet.row(i)[0].value))+"/"+ str(sheet.nrows-1) +" FETCHED")

    timenow = datetime.now()
    try:
        connectdevice = ConnectHandler(**device)
        print("CONNECTED TO DEVICE")
        fetch = connectdevice.send_command("show running-config", read_timeout=300)
        new_file_create = open(r"Output\Backup_"+sheet.row(i)[1].value+"-"+"-"+str(timenow.year)+"-"+str(timenow.month)+"-"+str(timenow.day)+"-"+str(timenow.hour)+"-"+str(timenow.minute)+"-"+str(timenow.second)+".cfg","w")
        copy_content = new_file_create.write(fetch)
        new_file_create.close()
        print("BACKUP GENERATED SUCCESSFULLY")
        connectdevice.disconnect()
        print("DISCONNECTED FROM DEVICE\n")
    except NetMikoTimeoutException:
        print("NetMikoTimeoutException - "+sheet.row(i)[1].value)
    except NetMikoAuthenticationException:
        print("NetMikoAuthenticationException - "+sheet.row(i)[1].value)

print("PROCESS FINISHED SUCCESSFULLY " + str(datetime.now()) + ", It took " + str(datetime.now()-timenow))
