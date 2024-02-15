dim count
count=0
name=inputbox("Enter a word","Finding vowels in a word","here")
length=len(name)
for i=1 to length
a=mid(name,i,1)
if a="A" or a="a" or a="E" or a="e" or a="i" or a="I" or a="o" or a="O" or a="u" or a="U" then
count=count+1
end if
next
if count>0 then
msgbox "The word "&name&" contain "&vblf&count&" vowels"
else
msgbox "The word "&name&" does not contain any vowels"
end if