randomize
val=87364
n=mid(rnd(val),3,1)
if n=1 then n=autogenerate
msgbox n
for j=1 to n
	randomize
	val=7489
	no=no&mid(rnd(val),3,1)
next
'msgbox no

'----------------
txt=inputbox("enter anythings")
msgbox crypt(txt)

function crypt(txt)
	for i=1 to len(txt)
	a= i mod n
	if a=0 then a=n-1
	crypt=crypt&chr(asc(mid(no,a,1)) xor asc(mid(txt,i,1)))
	next
end function
decode=crypt(txt)
msgbox crypt(decode)

function autogenerate
	randomize
	val=87364
	autogenerate=mid(rnd(val),3,1)
end function