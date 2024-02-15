n=inputbox("enter")
dr=1
nr=1
nn=n
r=inputbox("enter")
for a=1 to r-1
n=n-1
dr=dr*n
next
dr=dr*nn
for b=1 to r
nr=nr*b
next
msgbox dr/nr