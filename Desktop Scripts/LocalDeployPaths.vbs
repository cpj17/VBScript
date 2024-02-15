set KeyList = CreateObject("System.Collections.ArrayList")
set valueList = CreateObject("System.Collections.ArrayList")

'Here Set Default Value Index
DefaultValue = 2

'Here Add Keys
KeyList.add "1. WMS PORTTAL"
KeyList.add "2. TMS PORTTAL"
KeyList.add "3. ADMIN PORTTAL"
KeyList.add "4. D"


'Here Add Values Of The Keys
valueList.ADD "\\192.168.0.57\inetpub\PRDG11EOSV03WD\PRDG11EOSV03WB"
valueList.ADD "\\192.168.0.57\inetpub\PRDG11EOSV03WD\PRDG11EOSV03WB_TMS"
valueList.ADD "\\192.168.0.57\inetpub\PRDG11EOSV03WD\PRDG11EOSV03WB_ADMIN"
valueList.ADD "D:\"