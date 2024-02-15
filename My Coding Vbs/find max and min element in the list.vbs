set list=CreateObject("System.collections.arraylist")
n=inputbox("how many number you have to enter")
for i=1 to n
	num=inputbox("how many number you have to enter")
	list.add num
next
'msgbox join(list.toarray)

'sorting procedure
max=list(0)
min=list(0)
for j=0 to list.count-1
	if list(j)>max then
		max=list(j)
	end if
	if list(j)<min then
		min=list(j)
	end if
next
msgbox "the maximum number in the list is "&max&vblf&"the minimum number in the list is "&min