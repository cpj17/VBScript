msgbox "Think any number. I will assume that's x"
msgbox "Multiply into 2 which number you'll guess. like x*2"
msgbox "add 4 into which answer get from the previous step. like x*2 +4"
msgbox "multiply into 5 which answer get from the previous step. like x*2+4 *5"
msgbox "subract 20 in the previous which you get. like x*2+4*5 -20"
n=inputbox("Then type the final answer what will you get in the previous step")
if n mod 2=0 then
	msgbox "Your guess number is "&int(strreverse(n))&". Is this correct"
else
	msgbox "sorry to say something wrong in your calculation pls check then you will continue"
end if