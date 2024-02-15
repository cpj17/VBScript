a=inputbox("enter")
msgbox fact(a)

function fact(x)
ans=1
do
ans=ans*x
x=x-1
loop while x>0
fact=ans
end function