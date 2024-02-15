set obj=createobject("wscript.shell")
'random key generation
randomize
val=87364
n=mid(rnd(val),3,1)
if n=1 then n=autogenerate
msgbox n
for j=1 to n-1
	randomize
	val=7489
	no=no&mid(rnd(val),3,1)
next
function autogenerate
	randomize
	val=87364
	num=mid(rnd(val),3,1)
	if num=1 then call autogenerate
	autogenerate=num
end function
msgbox no
'**************************************  crypt function
function crypt(txt)
	for i=1 to len(txt)
	a= i mod n
	if a=0 then a=n-1
	crypt=crypt&chr(asc(mid(no,a,1)) xor asc(mid(txt,i,1)))
	next
end function
txt=inputbox("enter")
msgbox crypt(txt)