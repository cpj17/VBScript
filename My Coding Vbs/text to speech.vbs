set obj=createobject("sapi.spvoice")
set obj.voice = obj.getvoices.item(0)
txt=inputbox("Enter")
obj.speak txt