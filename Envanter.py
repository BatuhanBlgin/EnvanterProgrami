import psutil
import os
import wmi
import sqlite3 as sql
import win32com.client
import time


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



"""Disk Bilgilerini Alan Kod Bloğu"""
disk = psutil.disk_usage('/')
totalsize = disk.total / 1024 ** 3
totalsize = str(totalsize)
if totalsize[1] == ".":
    totalsize = str(totalsize[0:1] + " GB")
elif totalsize[2] == ".":
    totalsize =str(totalsize[0:2] + " GB")
elif totalsize[3] == ".":
    totalsize = str(totalsize[0:3] + " GB")
else:
    totalsize = str(totalsize[0:4] + " GB")

"""Kullanilan"""
disk = psutil.disk_usage('/')
usedsize = disk.used / 1024 ** 3
usedsize = str(usedsize)
if usedsize[1] == ".":
    usedsize = str(usedsize[0:1] + " GB")
elif usedsize[2] == ".":
    usedsize =str(usedsize[0:2] + " GB")
elif usedsize[3] == ".":
    usedsize = str(usedsize[0:3] + " GB")
else:
    usedsize = str(usedsize[0:4] + " GB")

"""Kullanilabilir"""
disk = psutil.disk_usage('/')
freesize = disk.free / 1024 ** 3
freesize = str(freesize)
if usedsize[1] == ".":
    freesize = str(freesize[0:1] + " GB")
elif usedsize[2] == ".":
    freesize =str(freesize[0:2] + " GB")
elif freesize[3] == ".":
    freesize = str(freesize[0:3] + " GB")
else:
    freesize = str(freesize[0:4] + " GB")



conn = sql.connect("kullanici.db")

cursor = conn.cursor()
cursor.execute(
    "CREATE TABLE IF NOT EXISTS kullanicilar (ID INTEGER PRIMARY KEY AUTOINCREMENT,kullanici_ad TEXT , bilgisayaradi TEXT,marka TXT,model TXT,serial TEXT UNIQUE, isletimsistemi TEXT,islemci TEXT, ram TEXT,takili_ram_Sayisi TEXT , ekrankarti TEXT, diskkapasite TEXT, diskkullanilan TEXT, diskkullanilabilir TEXT)")


cursor.execute("INSERT OR REPLACE INTO kullanicilar(kullanici_ad,bilgisayaradi,marka,model,serial,isletimsistemi,islemci,ram,takili_ram_Sayisi,ekrankarti,diskkapasite,diskkullanilan,diskkullanilabilir)  VALUES('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}');"
                           .format(my_system.UserName,os_info.CSName,my_system.Manufacturer,my_system.Model,str(pc_serial[13:]),os_info.Caption,proc_info.Name,round(system_ram),ram_sayac,gpu_info.VideoProcessor,totalsize,usedsize,freesize))

conn.commit()


print("İşlem Tamamlandı")
time.sleep(2)





