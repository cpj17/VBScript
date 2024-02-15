set KeyList = CreateObject("System.Collections.ArrayList")
set valueList = CreateObject("System.Collections.ArrayList")

'Here Set Default Value Index
DefaultValue = 2

'Here Add Keys
KeyList.add "WMS PORTAL"
KeyList.add "TMS PORTAL"
KeyList.add "ADMIN PORTAL"
KeyList.add "D"
KeyList.add "Local G!! WD"


'Here Add Values Of The Keys
valueList.ADD "\\192.168.0.57\inetpub\PRDG11EOSV03WD\PRDG11EOSV03WB"
valueList.ADD "\\192.168.0.57\inetpub\PRDG11EOSV03WD\PRDG11EOSV03WB_TMS"
valueList.ADD "\\192.168.0.57\inetpub\PRDG11EOSV03WD\PRDG11EOSV03WB_ADMIN"
valueList.ADD "D:\"
valueList.ADD """\\192.168.0.53\5000 Application Exe\PRDG11EOSV05APPWD_LOCAL"""