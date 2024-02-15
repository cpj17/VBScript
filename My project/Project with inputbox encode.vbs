word=inputbox("Entre a word")
if word="" then
msgbox "It is Empty So it will Quit"
wscript.quit
else
length=len(word)
for i=1 to length
a=mid(word,i,1)
if instr(a," ") then add=add&"000-"
if a="a" then add=add&"001-"
if a="b" then add=add&"002-"
if a="c" then add=add&"003-"
if a="d" then add=add&"004-"
if a="e" then add=add&"005-"
if a="f" then add=add&"006-"
if a="g" then add=add&"007-"
if a="h" then add=add&"008-"
if a="i" then add=add&"009-"
if a="j" then add=add&"010-"
if a="k" then add=add&"011-"
if a="l" then add=add&"012-"
if a="m" then add=add&"013-"
if a="n" then add=add&"014-"
if a="o" then add=add&"015-"
if a="p" then add=add&"016-"
if a="q" then add=add&"017-"
if a="r" then add=add&"018-"
if a="s" then add=add&"019-"
if a="t" then add=add&"020-"
if a="u" then add=add&"021-"
if a="v" then add=add&"022-"
if a="w" then add=add&"023-"
if a="x" then add=add&"024-"
if a="y" then add=add&"025-"
if a="z" then add=add&"026-"

'small letter conversion over

if a="A" then add=add&"101-"
if a="B" then add=add&"102-"
if a="C" then add=add&"103-"
if a="D" then add=add&"104-"
if a="E" then add=add&"105-"
if a="F" then add=add&"106-"
if a="G" then add=add&"107-"
if a="H" then add=add&"108-"
if a="I" then add=add&"109-"
if a="J" then add=add&"110-"
if a="K" then add=add&"111-"
if a="L" then add=add&"112-"
if a="M" then add=add&"113-"
if a="N" then add=add&"114-"
if a="O" then add=add&"115-"
if a="P" then add=add&"116-"
if a="Q" then add=add&"117-"
if a="R" then add=add&"118-"
if a="S" then add=add&"119-"
if a="T" then add=add&"120-"
if a="U" then add=add&"121-"
if a="V" then add=add&"122-"
if a="W" then add=add&"123-"
if a="X" then add=add&"124-"
if a="Y" then add=add&"125-"
if a="Z" then add=add&"126-"

'capital letter conversion over

next
msgbox left(add,len(add)-1),0,"SecretLang"
end if
set obj=createobject("wscript.shell")
wscript.sleep 200
obj.run"notepad"
wscript.sleep 200
obj.sendkeys "{tab}"
obj.sendkeys "{tab}"
obj.sendkeys "{tab}"
obj.sendkeys "{tab}"
obj.sendkeys "{tab}"
obj.sendkeys "{tab}"
obj.sendkeys "Decoding The Information"
obj.sendkeys "{enter}"
obj.sendkeys "{enter}"
obj.sendkeys left(add,len(add)-1)
ch=msgbox("Hit Enter To Close The Window",0+vbSystemModal,"Quit")
if ch=vbOK then
obj.sendkeys "%fx"
obj.sendkeys "{right}"
obj.sendkeys "{enter}"
end if