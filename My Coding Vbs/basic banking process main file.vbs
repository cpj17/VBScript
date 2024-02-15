'In this file associate with the file is - basic banking process associate file
set fso=createobject("Scripting.FileSystemObject")
set ofile=fso.opentextfile("D:\ProgramHub\VBScript\My Coding Vbs\basic banking process associate file.vbs",8)
set rvbf=fso.opentextfile("D:\ProgramHub\VBScript\My Coding Vbs\basic banking process associate file.vbs",1)
execute rvbf.readall
call mainmenu
sub mainmenu
	do
		ch=inputbox("1.Create Account"&vblf&"2.log in"&vbLf&"3.Exit","CPJ Bank","Here")
		select case ch
			case 1
				call create
			case 2
				call login
			case 3
				wscript.quit
			case else
				msgbox "Invalid choice pls press corrct choice"
		end select
	loop
end sub
sub create
	name=inputbox("enter your name"&vblf&"Press 9 to exit"&vbLf&"press 8 to back mainmenu","Sign Up","here")
	if name="9" then wscript.quit
	if name="8" then call mainmenu
	pass=inputbox("enter your password"&vblf&"Press 9 to exit"&vbLf&"press 8 to back mainmenu","Sign Up","here")
	if pass="9" then wscript.quit
	if pass="8" then call mainmenu
	Randomize			'acc no generation
	val=463137
	accno=Mid(Rnd(val),3,3)&Mid(Rnd(val),3,3)
	msgbox "Account created successfully..."&vbLf&name&" your acc no is "&accno
	qt=""""
	bal=00.00
	amt=msgbox("whether to add an initial amount in your account. It is optional",vbYesNo)
	if amt=vbyes then
		bal=inputbox(name&" please enter your initial amount")
		ofile.write vblf&"ballist.add "&bal
	end if
		ofile.write vbLf&"nlist.add "&qt&name&qt&vbLf
		ofile.write "aclist.add "&qt&accno&qt&vbLf
		ofile.write "plist.add "&qt&pass&qt&vbLf
		ofile.write "ballist.add "&"00.00"
	msgbox "Welcome "&name&vbLf&"  acc no : "&accno&vbLf&"balance : "&bal
	call mainmenu
end sub

sub login
	set rvbf=fso.opentextfile("D:\ProgramHub\VBScript\My Coding Vbs\basic banking process associate file.vbs",1)
	execute rvbf.readall
	name=inputbox("enter your name or account number "&vblf&"Press 9 to exit"&vbLf&"press 8 to back mainmenu","Login","here")
	if name="9" then wscript.quit
	if name="8" then call mainmenu
	pass=inputbox("enter your password"&vblf&"Press 9 to exit"&vbLf&"press 8 to back mainmenu","Login","here")
	if pass="9" then wscript.quit
	if pass="8" then call mainmenu
	for i=0 to nlist.count-1
		if (name=nlist(i) or name=aclist(i)) and pass=plist(i) then
			call options(nlist(i),i,aclist(i))
			ck=1
		end if
	next
	if ck=0 then
		msgbox "Invalid name or password"
		call login
	end if
end sub

sub options(name,i,accno)
	op=inputbox("Welcome "&name&vbLf&"  Options"&vbLf&"    1.check bank balance"&vbLf&"    2.deposit"&vbLf&"    3.withdrawal"&vbLf&"    4.Transfer"&vblf&"    Otherwise exit")
	select case op
		case 1
			ch=msgbox ("Your current balance is : "&ballist(i)&vblf&"press yes to go back otherwise goto main menu",vbYesNo)
			if ch=vbYes then call options(name,i,accno) else call mainmenu
		case 2
			dpamt=inputbox("Enter your amount")
			bal=ballist(i)+dpamt
			ofile.write vbLf&"ballist.item("&i&")="&bal
			execute rvbf.readall
			ch=msgbox ("Your current balance is : "&ballist(i)&vblf&"press yes to go back otherwise goto main menu",vbYesNo)
			if ch=vbYes then call options(name,i,accno) else call mainmenu
		case 3
			dpamt=inputbox("Enter your amount")
			bal=ballist(i)-dpamt
			ofile.write vbLf&"ballist.item("&i&")="&bal
			execute rvbf.readall
			ch=msgbox ("Your current balance is : "&ballist(i)&vblf&"press yes to go back otherwise goto main menu",vbYesNo)
			if ch=vbYes then call options(name,i,accno) else call mainmenu
		case 4
			tramt=inputbox("please enter your transfer amount..")
			tracno=inputbox("please be enter the account number which account you want transfer")
			for j=0 to aclist.count-1
				if tracno=aclist(j) then
					trbal=ballist(j)+tramt
					rembal=ballist(i)-tramt
					ofile.write vbLf&"ballist.item("&j&")="&trbal
					ofile.write vbLf&"ballist.item("&i&")="&rembal
				end if
			next
			execute rvbf.readall
			ch=msgbox ("Now your current account balance is :"&rembal&vblf&"press yes to go back otherwise goto main menu",vbYesNo)
			if ch=vbYes then call options(name,i,accno) else call mainmenu
		case else
			wscript.quit
	end select
end sub