import psutil
import os
import wmi
import sqlite3 as sql
import win32com.client


strComputer = "."
objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
ram_sayac = 0
for objItem in colItems:
    if objItem.BankLabel != None:
        ram_sayac +=1


computer = wmi.WMI()
my_system = computer.Win32_ComputerSystem()[0]
os_info = computer.Win32_OperatingSystem()[0]
proc_info = computer.Win32_Processor()[0]
gpu_info = computer.Win32_VideoController()[0]

os_name = os_info.Name.encode('utf-8').split(b"|")[0]
os_version = ' '.join([os_info.Version, os_info.BuildNumber])
system_ram = float(os_info.TotalVisibleMemorySize) / 1048576  # KB to GB
pc_serial = os.popen("wmic bios get serialnumber").read().replace("\n","")

disk = psutil.disk_usage('/')
size = disk.total / 1024 **3
size = str(size)
if size[1] == ".":
    totalsize = str(size[0:1]+" GB")
elif size[2] == ".":
    totalsize =str(size[0:2]+" GB")
elif size[3] == ".":
    totalsize = str(size[0:3]+" GB")
else:
    totalsize = str(size[0:4]+" GB")

marka_bilgi = my_system.Manufacturer
model_bilgi = my_system.Model
ram_sayisi = ram_sayac



conn = sql.connect("C:\\Users\\{}\\Desktop\kullanici.db".format(os_info.RegisteredUser))
cursor = conn.cursor()
cursor.execute(
    "CREATE TABLE IF NOT EXISTS kullanicilar (ID INTEGER PRIMARY KEY AUTOINCREMENT,kullanici_ad TEXT , bilgisayaradi TEXT,marka TXT,model TXT,serial TEXT, isletimsistemi TEXT,islemci TEXT, ram TEXT,takili_ram_Sayisi TEXT , ekrankarti TEXT, diskkapasite TEXT)")


cursor.execute("INSERT INTO kullanicilar(kullanici_ad,bilgisayaradi,marka,model,serial,isletimsistemi,islemci,ram,takili_ram_Sayisi,ekrankarti,diskkapasite)  VALUES('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}');"
                           .format(os_info.RegisteredUser,os_info.CSName,marka_bilgi,model_bilgi,str(pc_serial[13:]),os_info.Caption,proc_info.Name,round(system_ram),ram_sayisi,gpu_info.VideoProcessor,totalsize))
conn.commit()




