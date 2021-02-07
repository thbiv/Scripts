@ECHO OFF
net use z: /delete
net use z: \\thb-win7-vm\sysinfo$

systeminfo > C:\Temp.txt

findstr /C:"OS Name:" /C:"Original Install Date:" /C:"System Boot Time:" /C:"System Manufacturer:" /C:"System Model:" C:\Temp.txt > Z:\%computername%.txt

wmic bios get serialnumber > C:\sn.txt
Type C:\sn.txt >> z:\%computername%.txt

DEL C:\sn.txt
DEL C:\Temp.txt

net use Z: /delete