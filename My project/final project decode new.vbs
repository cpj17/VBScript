set obj=createobject("wscript.shell")
set fso=createobject("Scripting.FileSystemObject")
if fso.FileExists("C:\Users\PRAVEEN JEEVA\Desktop\decode.txt") then
	call Read_Decode
else
	call Write_Decode
end if


sub Read_Decode
	set filep=fso.getfile("C:\Users\PRAVEEN JEEVA\Desktop\decode.txt")
	if filep.size=0 then
		msgbox "this file was empty so you write or paste anything"
		call Write_Decode
	else
		set ofile=fso.opentextfile("C:\Users\PRAVEEN JEEVA\Desktop\decode.txt",1)
		content=split(ofile.readall,"-")
		for each i in content
			ncontent=ncontent&i
		next
		if isnumeric(ncontent) and instr(1,ncontent,"+") then
			msgbox "In this file contain alphabet or special character or not a correct format.The correct format like 001-002-003.So pls type or paste correct format",vbCritical
			call Write_Decode
		elseif not isnumeric(ncontent) then
			msgbox "In this file contain alphabet or special character or not a correct format.The correct format like 001-002-003.So pls type or paste correct format",vbCritical
			call Write_Decode
		elseif len(ncontent) mod 3<>0 then
			msgbox "In this file contain a not correct format or any number miss.The correct format like 001-002-003.So pls type or paste correct format",vbCritical
			call Write_Decode
		else
			call CrossCheck
		end if
	end if
end sub


sub Write_Decode
	fso.CreateTextFile("C:\Users\PRAVEEN JEEVA\Desktop\decode.txt")
	obj.run """C:\Users\PRAVEEN JEEVA\Desktop\decode.txt"""
	con=msgbox("we can start the conversion",vbYesNo+vbSystemModal)
	if con=vbyes then
		obj.sendkeys "%fs"
		wscript.sleep 100
		obj.sendkeys "%fx"
		call Read_Decode
	else
		obj.sendkeys "%fx"
		obj.sendkeys "{right}"
		obj.sendkeys "{enter}"
		wscript.quit
	end if
end sub


sub conversion
	'conversion is start
	set ofile=fso.opentextfile("C:\Users\PRAVEEN JEEVA\Desktop\decode.txt",1)
	do until ofile.atendofstream
		a=ofile.read(4)
		if a="888-" or a="888" then add=add&vbLf
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
	loop
	obj.run"notepad"
	wscript.sleep 200
	obj.sendkeys add
	con=msgbox("Hit Enter To Close the Window",vbOKOnly+vbSystemModal,"Quit")
	if con=vbOk then
		obj.sendkeys "%fx"
		obj.sendkeys "{right}"
		obj.sendkeys "{enter}"
		wscript.quit
	end if
end sub

sub CrossCheck
	set Decode_Data=CreateObject("System.Collections.ArrayList")
	set content_list=CreateObject("System.Collections.ArrayList")
	set uwd=CreateObject("System.Collections.ArrayList")
	set ofile=fso.OpenTextFile("C:\Users\PRAVEEN JEEVA\Desktop\decode.txt",1)
	do until ofile.atendofstream
		content_list.add ofile.read(4)
	loop
	
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
	
	Decode_Data.add "888"
	Decode_Data.add "888-"
	
	
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
		call Write_Decode
	end if
end sub