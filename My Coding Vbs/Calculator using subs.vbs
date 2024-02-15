sub data(a,b,op)
a=inputbox("Enter a",op)
b=inputbox("Eneter b",op)
end sub

sub data1(a,op)
a=inputbox("Enter a",op)
end sub

sub add(title)
call data(a,b,op)
msgbox a&" + "&b&" = "&eval(a)+eval(b),0,title
end sub

sub subract(title)
call data(a,b,op)
msgbox a&" - "&b&" = "&a-b,0,title
end sub

sub multi(title)
call data(a,b,op)
msgbox a&" * "&b&" = "&a*b,0,title
end sub

sub divi(title)
call data(a,b,op)
msgbox a&" / "&b&" = "&a/b,0,title
end sub

sub sqr(title)
call data1(a,op)
msgbox a*a,0,title
end sub

do until ch=6

ch=inputbox("1.Addition"&vblf&"2.Subraction"&vblf&"3.Multiplecation"&vbLf&"4.Divison"&vbLf&"5.Squre of a Value"&vblf&"6.Exit"&vbLf&"Enter your choice ")

if ch=1 then
op="Addition"
elseif ch=2 then
op="Subraction"
elseif ch=3 then
op="Multiplecation"
elseif ch=4 then
op="Divison"
elseif ch=5 then
op="squre of a value"
end if


select case ch
case 1
call add("Answer")
case 2
call subract("Answer")
case 3
call multi("Answer")
case 4
call divi("Answer")
case 5
call sqr("Answer")
case 6
wscript.quit
case else
msgbox"Pls enter a valid choice"
end select
loop