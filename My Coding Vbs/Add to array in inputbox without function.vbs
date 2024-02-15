a=array("Praveen","Jeeva")
do
value=inputbox(join(a,vbLf)&vbCr&vbCr&"Enter a Name")
redim preserve a(ubound(a)+1)
a(ubound(a))=value
loop until value=""
msgbox "You entered name is......"&vbLf&join(a,vbcr),0,"names"