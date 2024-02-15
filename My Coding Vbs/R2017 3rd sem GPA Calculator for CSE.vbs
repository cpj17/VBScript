dm=inputbox("Enter your GRADE of DM")
cddm=4
if dm="o" or dm="O" then
dm=10*4
elseif dm="a+" or dm="A+" then
dm=9*4
elseif dm="a" or dm="A" then
dm=8*4
elseif dm="b+" or dm="B+" then
dm=7*4
elseif dm="b" or dm="B" then
dm=6*4
else
dm=0
cddm=0
end if
dpsd=inputbox("Enter your GRADE of dpsd")
cddpsd=4
if dpsd="o" or dpsd="O" then
dpsd=10*4
elseif dpsd="a+" or dpsd="A+" then
dpsd=9*4
elseif dpsd="a" or dpsd="A" then
dpsd=8*4
elseif dpsd="b+" or dpsd="B+" then
dpsd=7*4
elseif dpsd="b" or dpsd="B" then
dpsd=6*4
else
dpsd=0
cddpsd=0
end if
ds=inputbox("Enter your GRADE of ds")
cdds=3
if ds="o" or ds="O" then
ds=10*3
elseif ds="a+" or ds="A+" then
ds=9*3
elseif ds="a" or ds="A" then
ds=8*3
elseif ds="b+" or ds="B+" then
ds=7*3
elseif ds="b" or ds="B" then
ds=6*3
else
ds=0
cdds=0
end if
ce=inputbox("Enter your GRADE of ce")
cdce=3
if ce="o" or ce="O" then
ce=10*3
elseif ce="a+" or ce="A+" then
ce=9*3
elseif ce="a" or ce="A" then
ce=8*3
elseif ce="b+" or ce="B+" then
ce=7*3
elseif ce="b" or ce="B" then
ce=6*3
else
ce=0
cdce=0
end if
oop=inputbox("Enter your GRADE of oop")
cdoop=3
if oop="o" or oop="O" then
oop=10*3
elseif oop="a+" or oop="A+" then
oop=9*3
elseif oop="a" or oop="A" then
oop=8*3
elseif oop="b+" or oop="B+" then
oop=7*3
elseif oop="b" or oop="B" then
oop=6*3
else
oop=0
cdoop=0
end if
dslab=inputbox("Enter your GRADE of dslab")
cddslab=2
if dslab="o" or dslab="O" then
dslab=10*2
elseif dslab="a+" or dslab="A+" then
dslab=9*2
elseif dslab="a" or dslab="A" then
dslab=8*2
elseif dslab="b+" or dslab="B+" then
dslab=7*2
elseif dslab="b" or dslab="B" then
dslab=6*2
else
dslab=0
cddslab=0
end if
ooplab=inputbox("Enter your GRADE of ooplab")
cdooplab=2
if ooplab="o" or ooplab="O" then
ooplab=10*2
elseif ooplab="a+" or ooplab="A+" then
ooplab=9*2
elseif ooplab="a" or ooplab="A" then
ooplab=8*2
elseif ooplab="b+" or ooplab="B+" then
ooplab=7*2
elseif ooplab="b" or ooplab="B" then
ooplab=6*2
else
ooplab=0
cdooplab=0
end if
dpsdlab=inputbox("Enter your GRADE of dpsdlab")
cddpsdlab=2
if dpsdlab="o" or dpsdlab="O" then
dpsdlab=10*2
elseif dpsdlab="a+" or dpsdlab="A+" then
dpsdlab=9*2
elseif dpsdlab="a" or dpsdlab="A" then
dpsdlab=8*2
elseif dpsdlab="b+" or dpsdlab="B+" then
dpsdlab=7*2
elseif dpsdlab="b" or dpsdlab="B" then
dpsdlab=6*2
else
dpsdlab=0
cddpsdlab=0
end if
englab=inputbox("Enter your GRADE of englab")
cdenglab=1
if englab="o" or englab="O" then
englab=10*1
elseif englab="a+" or englab="A+" then
englab=9*1
elseif englab="a" or englab="A" then
englab=8*1
elseif englab="b+" or englab="B+" then
englab=7*1
elseif englab="b" or englab="B" then
englab=6*1
else
englab=0
cdenglab=0
end if
dr=dm+oop+ds+dpsd+ce+dpsdlab+dslab+ooplab+englab
nr=cddm+cdoop+cdds+cddpsd+cdce+cddpsdlab+cddslab+cdooplab+cdenglab
gpa=dr/nr
msgbox formatnumber(gpa,2)