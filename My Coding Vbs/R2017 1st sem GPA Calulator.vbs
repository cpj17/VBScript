mat=inputbox("Enter your GRADE of Maths")
if mat="o" or mat="O" then
mat=10
cdmat=4
elseif mat="a+" or mat="A+" then
mat=9
cdmat=4
elseif mat="a" or mat="A" then
mat=8
cdmat=4
elseif mat="b+" or mat="B+" then
mat=7
cdmat=4
elseif mat="b" or mat="B" then
mat=6
cdmat=4
else
mat=5
cdmat=0
end if
eg=inputbox("Enter your GRADE of EG")
if eg="o" or eg="O" then
eg=10
cdeg=4
elseif eg="a+" or eg="A+" then
eg=9
cdeg=4
elseif eg="a" or eg="A" then
eg=8
cdeg=4
elseif eg="b+" or eg="B+" then
eg=7
cdeg=4
elseif eg="b" or eg="B" then
eg=6
cdeg=4
else
eg=5
cdeg=0
end if
eng=inputbox("Enter your GRADE of English")
if eng="o" or eng="O" then
eng=10
cdeng=4
elseif eng="a+" or eng="A+" then
eng=9
cdeng=4
elseif eng="a" or eng="A" then
eng=8
cdeng=4
elseif eng="b+" or eng="B+" then
eng=7
cdeng=4
elseif eng="b" or eng="B" then
eng=6
cdeng=4
else
eng=5
cdeng=0
end if
che=inputbox("Enter your GRADE of Chemistry")
if che="o" or che="O" then
che=10
cdche=3
elseif che="a+" or che="A+" then
che=9
cdche=3
elseif che="a" or che="A" then
che=8
cdche=3
elseif che="b+" or che="B+" then
che=7
cdche=3
elseif che="b" or che="B" then
che=6
cdche=3
else
che=5
cdche=0
end if
phy=inputbox("Enter your GRADE of Physics")
if phy="o" or phy="O" then
phy=10
cdphy=3
elseif phy="a+" or phy="A+" then
phy=9
cdphy=3
elseif phy="a" or phy="A" then
phy=8
cdphy=3
elseif phy="b+" or phy="B+" then
phy=7
cdphy=3
elseif phy="b" or phy="B" then
phy=6
cdphy=3
else
phy=5
cdphy=0
end if
psp=inputbox("Enter your GRADE of PSP")
if psp="o" or psp="O" then
psp=10
cdpsp=3
elseif psp="a+" or psp="A+" then
psp=9
cdpsp=3
elseif psp="a" or psp="A" then
psp=8
cdpsp=3
elseif psp="b+" or psp="B+" then
psp=7
cdpsp=3
elseif psp="b" or psp="B" then
psp=6
cdpsp=3
else
psp=5
cdpsp=0
end if
psplab=inputbox("Enter your GRADE of PSP Lab")
if psplab="o" or psplab="O" then
psplab=10
cdpsplab=2
elseif psplab="a+" or psplab="A+" then
psplab=9
cdpsplab=2
elseif psplab="a" or psplab="A" then
psplab=8
cdpsplab=2
elseif psplab="b+" or psplab="B+" then
psplab=7
cdpsplab=2
elseif psplab="b" or psplab="B" then
psplab=6
cdpsplab=2
else
psplab=5
cdpsplab=0
end if
bslab=inputbox("Enter your GRADE of BS Lab")
if bslab="o" or bslab="O" then
bslab=10
cdbslab=2
elseif bslab="a+" or bslab="A+" then
bslab=9
cdbslab=2
elseif bslab="a" or bslab="A" then
bslab=8
cdbslab=2
elseif bslab="b+" or bslab="B+" then
bslab=7
cdbslab=2
elseif bslab="b" or bslab="B" then
bslab=6
cdbslab=2
else
bslab=5
cdbslab=0
end if
dr=(cdmat*mat+cdeg*eg+cdeng*eng+cdche*che+cdpsp*psp+cdphy*phy+cdbslab*bslab+cdpsplab*psplab)
nr=cdche+cdphy+cdmat+cdpsp+cdeg+cdeng+cdpsplab+cdbslab
gpa=dr/nr
msgbox formatnumber(gpa,1)