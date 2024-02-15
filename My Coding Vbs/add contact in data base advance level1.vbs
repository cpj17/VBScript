'it is link in another file
'it has a data base that is add contact data base file

do
ch=inputbox("1.view"&vbLf&"2.add","","press 3 to exit")
select case ch
	case 1
		set fso=CreateObject("Scripting.FileSystemObject")
		set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\add contact data base file advance level1.vbs",1)
		Execute ofile.readall
		person=inputbox("enter name")
		msgbox contact.item(person)
	case 2
		call writecontact
	end select
loop until ch=3

sub writecontact
	set fso=CreateObject("Scripting.FileSystemObject")
	set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\add contact data base file advance level1.vbs",8)
	n=inputbox("how many contact you want to add")
	name=array()
	age=array()
	for i=0 to n-1
		redim preserve name(ubound(name)+1)
		name(ubound(name))=inputbox("enter " &i+1 & " contact name")
		redim preserve age(ubound(age)+1)
		age(ubound(age))=inputbox("enter "& name(i) &"'s age")
		quote=""""
		ofile.writeblanklines(1)
		ofile.write "contact.add"&" "&quote&name(i)&quote&","&age(i)
	next
end sub