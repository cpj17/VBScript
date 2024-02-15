n=inputbox("Enter No.of elements")
a=array()
for i=0 to n
	redim preserve a(ubound(a)+1)
	a(ubound(a))=inputbox("enter "&i&" element")
next
for i=0 to n
	element=element&a(i)&"  "
next
k=inputbox("search element")
for i=0 to n
	if k=a(i) then
		msgbox "given array is.."&vbLf&element&vbLf&k&" element position is "&i
		exit for
	end if
next	
if i=n then msgbox "given array is.."&vbLfk&" element not present in ana array"