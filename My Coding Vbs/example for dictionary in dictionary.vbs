set contact=CreateObject("Scripting.Dictionary")
contact.add "cpj",details("18","9715587522")
contact.add "jeeva",details("17","8667599658")
name=inputbox("which contact details you want")
msgbox contactdetails(name)

function details(age,phno)
	set details=CreateObject("Scripting.Dictionary")
	details.add "Age",age
	details.add "Phno",phno
end function

function contactdetails(person)
	contactdetails=person&vbLf
	allkeys=contact.item(person).keys()
	allitems=contact.item(person).items()
	for i=0 to contact.count-1
		contactdetails=contactdetails & "   " & allkeys(i) & " : " &allitems(i) & vblf
	next
end function