Call sum(x,y)
Function sum(x,y)
x=cint(inputbox("Enter a number A:"))
y=cint(inputbox("Enter a number B:"))
z=inputbox("Enter Your Choice:")

Select case z
case "+" msgbox x+y
case "-" msgbox x-y
case "*" msgbox x*y
case "/" msgbox x/y
End Select
End Function