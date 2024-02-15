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

ipdate=inputbox("Enter a date like"&vblf&"28-oct-2020")
difdate=datediff("d",now,ipdate)

coin=7565
dy=weekday(now)-1
for i=1 to difdate
	coin=coin+list(dy)
	dy=dy+1
	if dy=7 then dy=0
next

msgbox coin