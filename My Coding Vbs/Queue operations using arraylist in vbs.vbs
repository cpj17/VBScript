set queue=CreateObject("System.Collections.ArrayList")
cnt=array("1.Insert","2.Delete","3.Display","4.View FRONT","5.View REAR","6.Exit")
do
ch=inputbox(join(cnt,vblf))
select case ch
	case 1
		call insert
	case 2
		call delete
	case 3
		call disp
	case 4
		call dispfrnt
	case 5
		disprr
	case 6
		wscript.quit
	case else
		msgbox "Invalid choice"
end select
loop

sub insert
	n=inputbox("How many number of value you want to insert")
	for i=1 to n
		val=inputbox("Enter "&i&" Value")
		queue.add val
	next
end sub

sub disp
	if queue.count=0 then
		msgbox "Queue is empty"
	else
		msgbox "The elements in queue is.."&vblf&join(queue.toarray,"->")
	end if
end sub

sub delete
	if queue.count=0 then
		msgbox "Queue is empty"
	else
		queue.removeat 0
		msgbox "Deletion is completed"
	end if
end sub

sub dispfrnt
	if queue.count=0 then
		msgbox "Queue is empty"
	else
		msgbox "FRONT : "&queue(0)
	end if
end sub

sub disprr
	if queue.count=0 then
		msgbox "Queue is empty"
	else
		msgbox "REAR : "&queue(queue.count-1)
	end if
end sub
