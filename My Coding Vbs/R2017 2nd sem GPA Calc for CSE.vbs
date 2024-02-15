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
beee=inputbox("Enter your GRADE of BEEME")
if beee="o" or beee="O" then
beee=10
cdbeee=3
elseif beee="a+" or beee="A+" then
beee=9
cdbeee=3
elseif beee="a" or beee="A" then
beee=8
cdbeee=3
elseif beee="b+" or beee="B+" then
beee=7
cdbeee=3
elseif beee="b" or beee="B" then
beee=6
cdbeee=3
else
beee=5
cdbeee=0
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
evs=inputbox("Enter your GRADE of EVS")
if evs="o" or evs="O" then
evs=10
cdevs=3
elseif evs="a+" or evs="A+" then
evs=9
cdevs=3
elseif evs="a" or evs="A" then
evs=8
cdevs=3
elseif evs="b+" or evs="B+" then
evs=7
cdevs=3
elseif evs="b" or evs="B" then
evs=6
cdevs=3
else
evs=5
cdevs=0
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
cp=inputbox("Enter your GRADE of CP")
if cp="o" or cp="O" then
cp=10
cdcp=3
elseif cp="a+" or cp="A+" then
cp=9
cdcp=3
elseif cp="a" or cp="A" then
cp=8
cdcp=3
elseif cp="b+" or cp="B+" then
cp=7
cdcp=3
elseif cp="b" or cp="B" then
cp=6
cdcp=3
else
cp=5
cdcp=0
end if
cplab=inputbox("Enter your GRADE of CP Lab")
if cplab="o" or cplab="O" then
cplab=10
cdcplab=2
elseif cplab="a+" or cplab="A+" then
cplab=9
cdcplab=2
elseif cplab="a" or cplab="A" then
cplab=8
cdcplab=2
elseif cplab="b+" or cplab="B+" then
cplab=7
cdcplab=2
elseif cplab="b" or cplab="B" then
cplab=6
cdcplab=2
else
cplab=5
cdcplab=0
end if
eplab=inputbox("Enter your GRADE of EP Lab")
if eplab="o" or eplab="O" then
eplab=10
cdeplab=2
elseif eplab="a+" or eplab="A+" then
eplab=9
cdeplab=2
elseif eplab="a" or eplab="A" then
eplab=8
cdeplab=2
elseif eplab="b+" or eplab="B+" then
eplab=7
cdeplab=2
elseif eplab="b" or eplab="B" then
eplab=6
cdeplab=2
else
eplab=5
cdeplab=0
end if
dr=(cdmat*mat+cdbeee*beee+cdeng*eng+cdevs*evs+cdcp*cp+cdphy*phy+cdeplab*eplab+cdcplab*cplab)
nr=cdevs+cdphy+cdmat+cdcp+cdbeee+cdeng+cdcplab+cdeplab
gpa=dr/nr
msgbox formatnumber(gpa,1)