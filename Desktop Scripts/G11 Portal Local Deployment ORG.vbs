set obj = CreateObject("wscript.shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

path1 = "\\192.168.0.57\inetpub\PRDG11EOSV03WD\PRDG11EOSV03WB"
path2 = """\\192.168.0.57\Gurusoft\G11 Portal Backups\Backups"""
path3 = """\\192.168.0.57\Gurusoft\G11 Portal Backups\Inbox"""

obj.run path1
wscript.sleep 200
obj.run path2
wscript.sleep 200
obj.run path3