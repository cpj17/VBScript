set n1_list=createobject("System.Collections.Arraylist")
set n2_list=createobject("System.Collections.Arraylist")

n1=lcase(inputbox("Enter 1st name"))
n2=lcase(inputbox("Enter 2nd name"))

for i=1 to len(n1)
	x=mid(n1,i,1)
	if x<>" " then
		n1_list.add x
	end if
next

for i=1 to len(n2)
	y=mid(n2,i,1)
	if y<>" " then
		n2_list.add y
	end if
next

for i=0 to n1_list.count-1
	for j=0 to n2_list.count-1
		if(n1_list(i)=n2_list(j)) then
			a=a+1
			n2_list(j)="_"
			exit for
		end if
	next
next
count=(n1_list.count+n2_list.count)-(a*2)
msgbox count

if count=3 or count=5 or count=14 or count=16 or count=18 or count=21 or count=23 or count=32 or count=34 then
	msgbox "Friends"
elseif count=10 or count=19 or count=36 or count=37 or count=39 then
	msgbox "Lovers"
elseif count=8 or count=12 or count=13 or count=17 or count=28 or count=30 or count=35 then
	msgbox "Affection"
elseif count=6 or count=11 or count=15 or count=26 or count=31 or count=33 then
	msgbox "Marriage"
elseif count=2 or count=4 or count=7 or count=9 or count=20 or count=22 or count=24 or count=25 then
	msgbox "Enemy"
else
	msgbox "Sister"
end if
