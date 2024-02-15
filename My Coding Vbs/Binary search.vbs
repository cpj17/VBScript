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
beg=lbound(a)
dest=ubound(a)
do
	cen=(beg+dest)\2
	if a(cen)=k then
		msgbox "given array is"&vblf&element&vblf&k&" posisiton is "&cen
		pos=cen
	elseif a(cen)>k then
		dest=cen-1
	elseif a(cen)<k then
		beg=cen+1
	end if
loop until pos=cen