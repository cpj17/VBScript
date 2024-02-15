set fso=createobject("Scripting.FileSystemObject")
set ofile=fso.opentextfile("D:\ProgramHub\VBScript\My Coding Vbs\bank process just login and create assit code.vbs",8)
set rvbf=fso.opentextfile("D:\ProgramHub\VBScript\My Coding Vbs\bank process just login and create assit code.vbs",1)
execute rvbf.readall
sub mainmenu
do
ch=inputbox("1.Create Account"&vblf&"2.log in"&vbLf&"3.Exit","CPJ Bank","Here")
select case ch
case 1
call create
case 2
call login
case else
exit do
end select
loop until ch=9
end sub
call mainmenu
sub create
name=inputbox("enter your name")
pass=inputbox("enter your password")
Randomize
val=463137
accno=Mid(Rnd(val),3,3)&Mid(Rnd(val),3,3)
msgbox name&" your acc no is "&accno
qt=""""
ofile.write vbLf&"nlist.add "&qt&name&qt&vbLf
ofile.write "aclist.add "&qt&accno&qt&vbLf
ofile.write "plist.add "&qt&pass&qt
end sub

sub login
name=inputbox("enter your name or account number "&vblf&"Press 9 to exit"&vbLf&"press 8 to back mainmenu","Login","here")
if name="9" then wscript.quit
if name="8" then call mainmenu
pass=inputbox("enter your password"&vblf&"Press 9 to exit"&vbLf&"press 8 to back mainmenu","Login","here")
if pass="9" then wscript.quit
if pass="8" then call mainmenu
for i=0 to nlist.count-1
		if (name=nlist(i) or name=aclist(i)) and pass=plist(i) then
			msgbox "your name is :"&nlist(i)&vblf&"your account no :"&aclist(i)
			ck=1
		end if
	next
if ck=0 then
	msgbox "Invalid name or password"
	call login
end if
end sub