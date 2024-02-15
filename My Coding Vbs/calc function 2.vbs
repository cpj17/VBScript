call calc(x,y)
function calc(x,y)
x=cint(inputbox("Enter a number A:"))
y=cint(inputbox("Enter a number B:"))
z=inputbox("1.add"&vbcrlf&"2.subract"&vbcrlf&"3.multi"&vbcrlf&"4.divide")

select case z
case "1" msgbox x+y
case "2" msgbox x-y
case "3" msgbox x*y
case "4" msgbox x/y
end select
end function