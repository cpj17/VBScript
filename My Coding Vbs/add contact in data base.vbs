'it is link in another file
'it has a data base that is add contact data base file

do
ch=inputbox("1.view"&vbLf&"2.add","","press 3 to exit")
select case ch
	case 1
		set fso=CreateObject("Scripting.FileSystemObject")
		set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\add contact data base file.vbs",1)
		Execute ofile.readall
		person=inputbox("enter name")
		msgbox contact.item(person)
	case 2
		call writecontact
	end select
loop until ch=3

sub writecontact
	set fso=CreateObject("Scripting.FileSystemObject")
	set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\add contact data base file.vbs",8)
	name=inputbox("enter contact name")
	age=inputbox("enter age")
	quote=""""
	ofile.writeblanklines(1)
	ofile.write "contact.add"&" "&quote&name&quote&","&age
end sub