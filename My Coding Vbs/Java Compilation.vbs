name=inputbox("Enter a file name or class name if it is same","Java Compile....","name")
set a=createobject("wscript.shell")
a.run "cmd"
wscript.sleep 300
a.sendkeys "set path="
a.sendkeys """D:\java\bin"""
a.sendkeys "{enter}"
a.sendkeys "cd java"
a.sendkeys "{enter}"
a.sendkeys "javac "&name&".java"
a.sendkeys "{enter}"