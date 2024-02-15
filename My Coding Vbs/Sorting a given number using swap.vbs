set a=CreateObject("System.collections.arraylist")
n=inputbox("how many number you have to enter")
for i=1 to n
	num=inputbox("enter")
	a.add num
next
msgbox join(a.toarray)
for i=0 to n-2
	for j=i+1 to n-1
		if int(a(i))>int(a(j)) then
			temp=a(i)
			a(i)=a(j)
			a(j)=temp
		end if
	next
next
msgbox join(a.toarray)