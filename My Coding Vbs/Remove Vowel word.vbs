set list=createobject("system.collections.arraylist")
a=inputbox("Enter name")
for i=1 to len(a)'
	b=mid(a,i,1)
	if not(b="a" or b="A" or b="e" or b="i" or b="o" or b="u" or b="E" or b="I" or b="O" or b="U") then
		c=c&b
		count=count+1
	end if
next
vowel=len(a)-count
msgbox "Given Word"&vblf&a&vblf&"Vowel Words="&vowel&vblf&"After Removing Vowels"&vblf&c&vblf&"Non-Vowels="&count