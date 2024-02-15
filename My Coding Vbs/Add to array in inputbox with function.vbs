dim names
names=array("smith","warner","marsh")
returnvalue=newarray(names)

msgbox join(returnvalue,vblf)

function newarray(curntarray)
do
value=inputbox(join(curntarray,vblf),"Enter a Name")
redim preserve curntarray(ubound(curntarray)+1)
curntarray(ubound(curntarray))=value
loop until value=""
newarray=curntarray
end function