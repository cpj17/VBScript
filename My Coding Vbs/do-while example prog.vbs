set list=createobject("system.collections.arraylist")
i=1
do
msgbox "in do loop "&i
i=i+1
loop while i<5
msgbox "out of do loop"