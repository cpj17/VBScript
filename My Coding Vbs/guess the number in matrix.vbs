set a=createobject("system.collections.arraylist")
do
	randomize
	val=left(rnd*10,1)
	if not a.contains(val) and val<>0 then
		a.add val
	end if
loop while a.count<=8

'initial arranement
val=a(0)&vbtab&a(1)&vbtab&a(2)&vblf&a(3)&vbtab&a(4)&vbtab&a(5)&vblf&a(6)&vbtab&a(7)&vbtab&a(8)
call get_val

val=a(0)&vbtab&a(1)&vbtab&a(2)&vblf&a(3)&vbtab&a(4)&vbtab&a(5)&vblf&a(6)&vbtab&a(7)&vbtab&a(8)
call get_val

val=a(0)&vbtab&a(1)&vbtab&a(2)&vblf&a(3)&vbtab&a(4)&vbtab&a(5)&vblf&a(6)&vbtab&a(7)&vbtab&a(8)
msgbox "Your guess number is "&a(4)&". Is this correct"

sub get_val
	n=inputbox(val&vblf&"Guess any number from matrix and enter the row, which row contain your guess number")
	call get_order(n)
end sub

sub get_order(n)
	
	if n=1 then
		call row_1
	elseif n=2 then
		call row_2
	elseif n=3 then
		call row_3
	else
		msgbox "Invalid choice"
		call get_val
	end if
end sub

sub row_1
	'matrix order will be change in this section ex 1 2 3 into 3 1 2
	call swap(0,1,2)
	call swap(3,4,5)
	call swap(6,7,8)
	'matrix order is done
	
	'order changeed matrix will be transpoed in row by column or column by row ex  3 1 2
	call swap_trans_val(1,3)													  '3
	call swap_trans_val(2,6)													  '1
	call swap_trans_val(5,7)													  '2
end sub

sub row_2
	call swap_trans_val(1,3)
	call swap_trans_val(2,6)
	call swap_trans_val(5,7)
end sub

sub row_3
	call swap(2,1,0)
	call swap(5,4,3)
	call swap(8,7,6)
	
	call swap_trans_val(1,3)
	call swap_trans_val(2,6)
	call swap_trans_val(5,7)
end sub
sub swap(m,n,k)
	t1=a(m)
	t2=a(n)
	a(m)=a(k)
	a(n)=t1
	a(k)=t2
end sub

sub swap_trans_val(m,n)
	t=a(m)
	a(m)=a(n)
	a(n)=t
end sub