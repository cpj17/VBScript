do while choice<>5
choice=inputbox("1.Add"&vbNewLine&"2.Subraction"&vbNewLine&"3.Multiplication"&vbNewLine&"4.Divison"&vbNewLine&"5.Exit"&vbNewLine&"Enter your choice","Calculator","here")
select case choice
case "1"
a=cdbl(inputbox("Enter A","Addition","here"))
b=cdbl(inputbox("Enter B","Addition","here"))
msgbox a&" + "&b&" = "&a+b
case "2"
a=cdbl(inputbox("Enter A","Subraction","here"))
b=cdbl(inputbox("Enter B","Subraction","here"))
msgbox a&" - "&b&" = "&a-b
case "3"
a=cdbl(inputbox("Enter A","Multiplication","here"))
b=cdbl(inputbox("Enter B","Multiplication","here"))
msgbox a&" * "&b&" = "&a*b
case "4"
a=cdbl(inputbox("Enter Divider","Divison","here"))
b=cdbl(inputbox("Enter Dividend","Divison","here"))
msgbox a&" / "&b&" = "&a/b
end select
loop