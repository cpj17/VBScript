on error resume next
name=array()
set contact=CreateObject("Scripting.Dictionary")
n=inputbox("how many contact you want to add")
for i=0 to n-1
	redim preserve name(ubound(name)+1)
	name(ubound(name))=inputbox("enter "&i+1&" contact name")
	contact.add name(ubound(name)),details(inputbox("enter "& name(ubound(name)) & "'s age"),(inputbox("enter "& name(ubound(name)) & "'s phno")))
next

function details(age,phno)
	set details=CreateObject("Scripting.Dictionary")
	details.add "Age",age
	details.add "Phno",phno
end function

function contactdetails(person)
	contactdetails=ucase(person)&vblf
	allkeys=contact.item(person).keys()
	allitems=contact.item(person).items()
	for i=0 to ubound(allitems)
		contactdetails=contactdetails & "  " & allkeys(i) &"   : "&allitems(i)&vbLf
	next
end function

msgbox contactdetails(inputbox("which contact details you want"))