set list=createobject("system.collections.arraylist")

day1=10+135
day2=20+135
day3=40+135
day4=25+135
day5=30+135
day6=35+135
day7=60+135

list.add day1
list.add day2
list.add day3
list.add day4
list.add day5
list.add day6
list.add day7

do
ipch=inputbox("1.Calculate day by day"&vblf&"2.Calculate with given date"&vblf&"3.Calculate with particular day in a number"&vblf&"4.Exit")
if ipch="1" or ipch="2" or ipch="3" or ipch="4" then
	if ipch=1 then
		call daybyday
	elseif ipch=2 then
		call calcwdate
	elseif ipch=3 then
		call calcwnum
	elseif ipch=4 then
		wscript.quit
	end if
else
	msgbox "pls enter correct choice"
end if
loop

sub daybyday
	coin=7565
	dy=2
	ch=msgbox ("Your last entered coin is 7565 data 17-March-2020 is that ok for you",vbYesNo)
	if ch=vbno then
		coin=inputbox("pls enter a coins currently you have")
		dy=weekday(now)-1
	end if
	i=1
	do
		m_date=dateadd("d",i,date)
		ch=msgbox (m_date&vblf&coin+list(dy),vbYesNo)
		coin=coin+list(dy)
		dy=dy+1
		i=i+1
		if dy=7 then dy=0
	loop until ch=vbno
end sub

sub calcwdate
	ipdate=inputbox("Enter a date like"&vblf&"28-oct-2020")
	difdate=datediff("d",now,ipdate)
	coin=7565
	dy=2
	ch=msgbox ("Your last entered coin is 7565 data 17-March-2020 is that ok for you",vbYesNo)
	if ch=vbno then
		coin=inputbox("pls enter a coins currently you have")
		dy=weekday(now)-1
	end if
	for i=1 to difdate
		coin=coin+list(dy)
		dy=dy+1
		if dy=7 then dy=0
	next
	msgbox coin&vblf&"days "&difdate
end sub

sub calcwnum
	num=inputbox("How many days want to calculate")
	coin=7565
	dy=2
	ch=msgbox ("Your last entered coin is 7565 data 17-March-2020 is that ok for you",vbYesNo)
	if ch=vbno then
		coin=inputbox("pls enter a coins currently you have")
		dy=weekday(now)-1
	end if
	for i=1 to num
		coin=coin+list(dy)
		dy=dy+1
		if dy=7 then dy=0
	next
	msgbox coin
end sub
