n=inputbox("Which Number table you want")
r=inputbox("And Range")
for i=1 to r
ans=ans&i&" "&"X "&n&" = "&i*n&vblf
next
msgbox ans