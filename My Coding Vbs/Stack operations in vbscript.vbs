stack=array()
args=array("1.Insert","2.Delete","3.Display","4.View Top Element","5.Exit")
do
ch=inputbox(join(args,vblf))
select case ch
	case 1
		call insert
	case 2
		call delete
	case 3
		if ubound(stack)=-1 then
			msgbox "Stack is empty"
		else
			msgbox "The Elements in stack is.."& vblf &display
		end if
	case 4
		call disptop
	case 5
		wscript.quit
	case else
		msgbox "Invalid Choice"
end select
loop

sub insert
	n=inputbox("How many element you want to insert")
		for i=1 to n
			val=inputbox("Enter "& i &" Element")
				redim preserve stack(ubound(stack)+1)
				stack(ubound(stack))=val
		next
	msgbox "Insertion is completed"
end sub

function display
	for i=0 to ubound(stack)
		elmnt=elmnt&stack(i)&"->"
	next
	display=left(elmnt,len(elmnt)-2)
end function

sub delete
	if ubound(stack)=-1 then
		msgbox "Stack is Empty"
	else
		redim preserve stack(ubound(stack)-1)
		msgbox "Deletion is completed"
	end if
end sub

sub disptop
	if ubound(stack)=-1 then
		msgbox "The stack has no element that means the stack is empty"
	else
		msgbox "The Top Element of Stack is : "&stack(ubound(stack))
	end if
end sub