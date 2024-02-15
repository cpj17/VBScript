fib=array(0,1,1)
n=inputbox("no.of fibonosi you want")
for i=1 to n-3
	redim preserve fib(ubound(fib)+1)
	fib(ubound(fib))=fib(ubound(fib)-2)+fib(ubound(fib)-1)
next
msgbox join(fib)