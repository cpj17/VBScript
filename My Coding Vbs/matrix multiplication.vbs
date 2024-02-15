dim a()
dim b()
dim c()
n=inputbox("enter row")
m=inputbox("enter column")
redim a(n,m)
redim b(n,m)
redim c(n,m)

sub get_matrix(mat)
	for i=0 to n
		for j=0 to m
			mat(i,j)=inputbox("enter "&i&" "&j&" element")
		next
	next
end sub

msgbox "enter elements of A matrix"
call get_matrix(a)
msgbox "enter elements of B matrix"
call get_matrix(b)

for i=0 to n
	for j=0 to m
		c(i,j)=0
		for k=0 to m
			c(i,j)=int(c(i,j))+int(a(i,k))*int(b(k,j))
		next
	next
next

for i=0 to n
	for j=0 to m
		element=element&c(i,j)&vbTab
	next
	element=element&vbLf
next

msgbox element