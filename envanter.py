import os
import wmi
import sqlite3 as sql
import win32com.client
import time

"bilgisayarda kaç tane ram var bulan kod"
strComputer = "."
objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
ram_sayac = 0
for objItem in colItems:
    if objItem.BankLabel != None:
        ram_sayac +=1

"bilgisayar bilgilerini toplayan kod"
computer = wmi.WMI()
my_system = computer.Win32_ComputerSystem()[0]
os_info = computer.Win32_OperatingSystem()[0]
proc_info = computer.Win32_Processor()[0]
gpu_info = computer.Win32_VideoController()[0]
my_Pyhsical = computer.Win32_PhysicalMemory()[0]

"""ram ddr bulan kod"""
def system_ramddr (ram):
    memory_types = {"0" :"Unknown","1":"Other","2":"DRAM","3":"Synchronous DRAM","4":"Cache DRAM","5":"EDO","6":"EDRAM","7":"VRAM","8":"SRAM","9":"RAM",
                    "10":"ROM","11":"Flash","12":"EEPROM","13":"FEPROM","14":"EPROM","15":"CDRAM","16":"3DRAM","17":"SDRAM","18":"SGRAM","19":"RDRAM","20":"DDR"
                    ,"21":"DDR2","22":"DDR2 FB-DIMM","24":"DDR3","25":"FBD2","26":"DDR4"}

    if ram in memory_types:
        x = memory_types[ram]
        return x


ramddr = system_ramddr(str(my_Pyhsical.SMBIOSMemoryType))


system_ram = float(os_info.TotalVisibleMemorySize) / 1048576  # KB to GB


pc_serial = os.popen("wmic bios get serialnumber").read().replace("\n","")

disktotal = os.popen(f'powershell.exe -Command "Get-PhysicalDisk | Select Size"').read().split()
totallist = []
del disktotal[:2]
for a in disktotal:
    a = int(a) // (2 ** 30)
    totallist.append(a)

totallist = str(totallist)
totallist = totallist.replace('[', ' ')
totallist = totallist.replace(']', ' ')
totallist = totallist.replace(',', ' ')

diskused = os.popen(f'powershell.exe -Command "Get-PSDrive | Select Used"').read().split()
usedlist = []
del diskused[:2]
diskused.reverse()
for b in diskused:
    b = int(b) // (2 ** 30)
    usedlist.append(b)

usedlist = str(usedlist)
usedlist = usedlist.replace('[', ' ')
usedlist = usedlist.replace(']', ' ')
usedlist = usedlist.replace(',', ' ')

diskfree = os.popen(f'powershell.exe -Command "Get-PSDrive | Select Free"').read().split()
freelist = []
del diskfree[:2]
diskfree.reverse()
for c in diskfree:
    c = int(c) // (2 ** 30)
    freelist.append(c)

freelist = str(freelist)
freelist = freelist.replace('[', ' ')
freelist = freelist.replace(']', ' ')
freelist = freelist.replace(',', ' ')


"""Bilgisayarda takılı olan diskin tipini ve adını gösteren kod"""
diskbilgi = os.popen(f'powershell.exe -Command "Get-PhysicalDisk | Select FriendlyName"').read()
disktipi = os.popen(f'powershell.exe -Command "Get-PhysicalDisk | Select MediaType"').read()

disktipi = disktipi.replace('MediaType',' ')
disktipi = disktipi.replace('-',' ')

diskbilgi = diskbilgi.replace('FriendlyName',' ')
diskbilgi = diskbilgi.replace('-',' ')




conn = sql.connect("kullanici.db")

cursor = conn.cursor()
cursor.execute(
    "CREATE TABLE IF NOT EXISTS kullanicilar (ID INTEGER PRIMARY KEY AUTOINCREMENT,kullanici_ad TEXT , bilgisayaradi TEXT,marka TXT,model TXT,serial TEXT UNIQUE, isletimsistemi TEXT,islemci TEXT, ram TEXT,takili_ram_Sayisi TEXT,ramddr TEXT , ekrankarti TEXT, diskbilgi TEXT,disktipi TEXT ,diskkapasite TEXT, diskkullanilan TEXT, diskkullanilabilir TEXT)")


cursor.execute("INSERT OR REPLACE INTO kullanicilar(kullanici_ad,bilgisayaradi,marka,model,serial,isletimsistemi,islemci,ram,takili_ram_Sayisi,ramddr,ekrankarti,diskbilgi,disktipi,diskkapasite,diskkullanilan,diskkullanilabilir)  VALUES('{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}','{}');"
                           .format(my_system.UserName,os_info.CSName,my_system.Manufacturer,my_system.Model,str(pc_serial[13:]),os_info.Caption,proc_info.Name,round(system_ram),ram_sayac,ramddr,gpu_info.VideoProcessor,diskbilgi.strip(),disktipi.strip(),str(totallist).strip(),str(usedlist).strip(),str(freelist).strip()))

conn.commit()



print("İşlem Tamamlandı")
time.sleep(2)
