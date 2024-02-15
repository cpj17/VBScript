dm=inputbox("Enter DM mark")
dpsd=inputbox(" Enter DPSD mark")
ds=inputbox("Enter DS mark")
oop=inputbox("Enter OOP mark")
ce=inputbox("Entre CE mark")
if dm=100 then ndm=10 else ndm=left(dm,1)+1
if dpsd=100 then ndpsd=10 else ndpsd=left(dpsd,1)+1
if ds=100 then nds=10 else nds=left(ds,1)+1
if ce=100 then nce=10 else nce=left(ce,1)+1
if oop=100 then noop=10 else noop=left(oop,1)+1
gpa=(4*ndm+4*ndpsd+3*nds+3*noop+3*nce)/17
msgbox "Your GPA is "&left(gpa,5),0,"GPA"