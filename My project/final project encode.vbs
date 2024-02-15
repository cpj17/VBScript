set obj=createobject("wscript.shell")
set fso=createobject("Scripting.FileSystemObject")
if fso.FileExists ("C:\Users\PRAVEEN JEEVA\Desktop\decode.txt") then 'if decode file is already exixts we don't over write
	fso.DeleteFile("C:\Users\PRAVEEN JEEVA\Desktop\decode.txt")      'so we can delete
end if
if fso.FileExists("C:\Users\PRAVEEN JEEVA\Desktop\encode.txt") then  'to check if the encode file exists or not
	call Read_Encode   												 'if the file is exists go to the Read_Encode sub
else
	call Write_Encode												 'if the file does not exists go to Write_Encode
end if


sub Read_Encode
	set filep=fso.getfile("C:\Users\PRAVEEN JEEVA\Desktop\encode.txt")								'used to get the file property
	if filep.size=0 then											'to check file is empty or not
		msgbox "This file is empty pls type anything"				'if size = 0 then the file is empty then it will display the file is empty
		call Write_Encode											'if the file is empty we do not start the conversion so we can write the file
	else															'if file is not empty so we can start the conversion
		set ofile=fso.opentextfile("C:\Users\PRAVEEN JEEVA\Desktop\encode.txt",1) 'used to open a file
		do until ofile.atendofstream
		
			'conversion is start
			
			a=ofile.read(1)											'read the character in one by one
			if a=vbLf then add=add&"888-"
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
		loop
		obj.run"notepad"
		wscript.sleep 200
		obj.sendkeys left(add,len(add)-1)
		wscript.sleep 500
		obj.sendkeys "%fs"
		wscript.sleep 400
		obj.sendkeys "decode"
		wscript.sleep 200
		obj.sendkeys "{enter}"
		wscript.sleep 100
		con=msgbox("Hit Enter To Close The Window",vbOK+vbSystemModal,"Quit")
		if con=vbOK then
			obj.sendkeys "%fx"
			obj.sendkeys "{right}"
			obj.sendkeys "{enter}"
		end if
	end if
end sub

sub Write_Encode
	fso.createtextfile("C:\Users\PRAVEEN JEEVA\Desktop\encode.txt")
	obj.run """C:\Users\PRAVEEN JEEVA\Desktop\encode.txt"""
	con=msgbox("can i start the conversion",vbYesNo+vbSystemModal+vbExclamation)
	if con=vbYes then
		obj.sendkeys "%fs"			'save the file
		wscript.sleep 100
		obj.sendkeys "%fx"			'exit the current text file
		call Read_Encode			'writing is over then we can read the file so call Read_Encode sub
	else
		obj.sendkeys "%fx"			'used to close the current text file
		obj.sendkeys "{right}"		'if any changes can made by user then it displays the msg save or not
		obj.sendkeys "{enter}"		'so we can do not save the text file by hiting the enter
		wscript.quit				'here the total code was terminated
	end if
end sub