x=array(10,12,9,3,18,7)
	
for i=ubound(x)-1 to 0 step -1
	for j=0 to i
		if x(j)>x(j+1) then
			temp=x(j+1)
			x(j+1)=x(j)
			x(j)=temp
		end if
	next
next

for each k in x
n=n&k&"  "
next
msgbox n