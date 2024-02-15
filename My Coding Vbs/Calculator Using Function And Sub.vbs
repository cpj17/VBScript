function add
call data(a,b)
add=int(a)+int(b)
end function

function subract
call data(a,b)
subract=a-b
end function

function multi
call data(a,b)
multi=a*b
end function

function divi
call data(a,b)
divi=a/b
end function

sub data(x,y)
x=inputbox("Enter x")
y=inputbox("Enter y")
end sub

op=array("1.Addition","2.Subraction","3.Multiplication","4.Division")
ch=inputbox(join(op,vbLf),"Calculator")
select case ch

case 1
msgbox "Ans = "&add
case 2
msgbox "Ans = "&subract
case 3
msgbox "Ans = "&multi
case 4
msgbox "Ans = "&divi
case else
msgbox "your choice is wrong so this process is quit"
wscript.quit

end select