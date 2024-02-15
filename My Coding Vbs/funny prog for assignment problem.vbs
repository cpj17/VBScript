set a=createobject("system.collections.arraylist")
n=inputbox("enter")
for i=1 to len(n)
	a.add mid(n,i,1)
next
init=a(0)&a(1)&a(2)&a(3)&vblf
for j=1 to 3
	t=a(2)
	a(2)=a(3)
	a(3)=t
	val=val&a(0)&a(1)&a(2)&a(3)&vblf
	
	t=a(1)
	a(1)=a(3)
	a(3)=t
	val=val&a(0)&a(1)&a(2)&a(3)&vblf
next

msgbox init&mid(val,1,len(val)-5)