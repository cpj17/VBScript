set Decode_Data=CreateObject("System.Collections.ArrayList")
set content_list=CreateObject("System.Collections.ArrayList")
set uwd=CreateObject("System.Collections.ArrayList")

call Write(word)
sub Write(word)
	word=inputbox("enter")
	'msgbox word
	if word="" then
		msgbox "It is empty"
	else
		call Check_1(word)
	end if
end sub

sub Check_1(word)
	for each content in Split(word,"-")
		ncontent=ncontent&content
		count=count+1
	next
	if Not IsNumeric(ncontent) then
		msgbox"It is contain Some letters or Special characters. So pls write or paste correct content"
		call Write(word)
	elseif IsNumeric(ncontent) and instr(ncontent,"+") then
		msgbox"It is contain Some letters or Special characters. So pls write or paste correct content"
		call Write(word)
	elseif len(ncontent) mod 3<>0 then
		msgbox "Missed any content"
		call Write(word)
	else
		call Check_2(count)
	end if
end sub	

sub Check_2(count)
	for i=0 to 26
		if len(i)=1 then
			Decode_Data.add "00"&i
			Decode_Data.add "00"&i&"-"
			Decode_Data.add "10"&i
			Decode_Data.add "10"&i&"-"
		else
			Decode_Data.add "0"&i
			Decode_Data.add "0"&i&"-"
			Decode_Data.add "1"&i
			Decode_Data.add "1"&i&"-"
		end if
	next
	for i=1 to count*4 step 4
		content_list.add mid(word,i,4)
	next
	for i=0 to content_list.count-1
		if Decode_Data.contains(content_list(i)) then
			count=count+1
		else
			uwd.add content_list(i)
			newcount=newcount+1
		end if
	next
	if uwd.count=0 then
		call conversion
	else
		msgbox "In this file contain "&newcount&" unwanted data such as "&left(join(uwd.toarray,""),(newcount*4)-1)&" so pls paste or write correct data"
		CreateObject("Wscript.shell").run """C:\Users\PRAVEEN JEEVA\Desktop\Project with inputbox decode.vbs"""
	end if
end sub
sub conversion
length=len(word)
for i=1 to length
a=mid(word,i,4)
if a="000-" or a="000" then add=add&" "
if a="001-" or a="001" then add=add&"a"
if a="002-" or a="002" then add=add&"b"
if a="003-" or a="003" then add=add&"c"
if a="004-" or a="004" then add=add&"d"
if a="005-" or a="005" then add=add&"e"
if a="006-" or a="006" then add=add&"f"
if a="007-" or a="007" then add=add&"g"
if a="008-" or a="008" then add=add&"h"
if a="009-" or a="009" then add=add&"i"
if a="010-" or a="010" then add=add&"j"
if a="011-" or a="011" then add=add&"k"
if a="012-" or a="012" then add=add&"l"
if a="013-" or a="013" then add=add&"m"
if a="014-" or a="014" then add=add&"n"
if a="015-" or a="015" then add=add&"o"
if a="016-" or a="016" then add=add&"p"
if a="017-" or a="017" then add=add&"q"
if a="018-" or a="018" then add=add&"r"
if a="019-" or a="019" then add=add&"s"
if a="020-" or a="020" then add=add&"t"
if a="021-" or a="021" then add=add&"u"
if a="022-" or a="022" then add=add&"v"
if a="023-" or a="023" then add=add&"w"
if a="024-" or a="024" then add=add&"x"
if a="025-" or a="025" then add=add&"y"
if a="026-" or a="026" then add=add&"z"

'small letter conversion over

if a="101-" or a="101" then add=add&"A"
if a="102-" or a="102" then add=add&"B"
if a="103-" or a="103" then add=add&"C"
if a="104-" or a="104" then add=add&"D"
if a="105-" or a="105" then add=add&"E"
if a="106-" or a="106" then add=add&"F"
if a="107-" or a="107" then add=add&"G"
if a="108-" or a="108" then add=add&"H"
if a="109-" or a="109" then add=add&"I"
if a="110-" or a="110" then add=add&"J"
if a="111-" or a="111" then add=add&"K"
if a="112-" or a="112" then add=add&"L"
if a="113-" or a="113" then add=add&"M"
if a="114-" or a="114" then add=add&"N"
if a="115-" or a="115" then add=add&"O"
if a="116-" or a="116" then add=add&"P"
if a="117-" or a="117" then add=add&"Q"
if a="118-" or a="118" then add=add&"R"
if a="119-" or a="119" then add=add&"S"
if a="120-" or a="120" then add=add&"T"
if a="121-" or a="121" then add=add&"U"
if a="122-" or a="122" then add=add&"V"
if a="123-" or a="123" then add=add&"W"
if a="124-" or a="124" then add=add&"X"
if a="125-" or a="125" then add=add&"Y"
if a="126-" or a="126" then add=add&"Z"

'capital letter conversion over

next
msgbox add
end sub