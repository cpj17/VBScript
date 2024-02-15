set fso=createobject("Scripting.FileSystemObject")
sub read(c)
set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\test.txt",1)
c=ofile.readall
end sub
call read(c)
msgbox "Before"&vblf&c
call wrt

sub wrt
set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\test.txt",8)

do
nc=inputbox(c,"Enter")
ofile.write(nc)
call read(c)
msgbox "after"&vblf&c
loop until nc=""
end sub