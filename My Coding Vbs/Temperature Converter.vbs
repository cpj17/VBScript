chd=array("1.Fahrenheit to Celsius","2.Celsius To Fahrenheit","Enter Your Choice")
ch=inputbox(join(chd,vblf),"Temperature Converter","here")

select case ch
	case 1
		f=inputbox("Enter Heat In Fahrenheit")
		c=(f-32)*5/9
		msgbox f&" Fahrenheit = "&c&" Celsius"
	case 2
		c=inputbox("Enter Heat In Celsius")
		f=(c*9/5)+32
		msgbox c&" Celsius = "&f&" Fahrenheit"
end select