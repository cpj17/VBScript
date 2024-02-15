set obj = CreateObject("Wscript.shell")

obj.sendkeys "{ }"
wscript.sleep 200
obj.sendkeys "%{f1}"