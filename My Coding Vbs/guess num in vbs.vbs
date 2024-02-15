name=inputbox("Enter Your Name","CPJ","name")
msgbox(name+" this is guess the integer number game")
msgbox"i have already set a number"
do
a=inputbox("Enter your guess number","CPJ","number")
if(a>17) then msgbox name+" your guess is high" else if(a<17) then msgbox name+" your guess is low" else msgbox"cool your guess is correct "+name
loop until(a=17)
msgbox"thanks for playing this game"