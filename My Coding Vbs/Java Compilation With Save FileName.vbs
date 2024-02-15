set fso=CreateObject("Scripting.FileSystemObject")
set a=createobject("wscript.shell")

sub read(fn)
set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\java\getfilename.txt",1)
fn=ofile.ReadAll
end sub
call read(fn)

nf=inputbox("If File Name is Same Press Enetr","title",fn)

if nf=fn then
a.run "cmd"
wscript.sleep 300
a.sendkeys "set path="
a.sendkeys """D:\java\bin"""
a.sendkeys "{enter}"
a.sendkeys "cd java"
a.sendkeys "{enter}"
a.sendkeys "javac "&nf&".java"
a.sendkeys "{enter}"
else
set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\java\getfilename.txt",2)
ofile.write(nf)
call read(fn)
a.run "cmd"
wscript.sleep 200
a.sendkeys "set path="
a.sendkeys """D:\java\bin"""
a.sendkeys "{enter}"
a.sendkeys "cd java"
a.sendkeys "{enter}"
a.sendkeys "javac "&nf&".java"
a.sendkeys "{enter}"
end if