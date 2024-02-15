txt=inputbox("enter anythings")
msgbox crypt(txt)

function crypt(txt)
	for i=1 to len(txt)
	a= i mod 5
	if a=0 then a=4
	crypt=crypt&chr(asc(mid(5849,a,1)) xor asc(mid(txt,i,1)))
	next
end function
decode=crypt(txt)
msgbox crypt(decode)