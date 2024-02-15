set list=createobject("system.collections.arraylist")

day1=10+160
day2=20+160
day3=40+160
day4=25+160
day5=30+160
day6=35+160
day7=60+160

list.add day1
list.add day2
list.add day3
list.add day4
list.add day5
list.add day6
list.add day7

coin=inputbox("pls enter a coins currently ")
dy=weekday(now)-1
i=1
dy_cnt=1
do
	m_date=dateadd("d",i,date)
	ch=msgbox (m_date&vblf&coin+list(dy)&vblf&"Day "&dy_cnt,vbYesNo)
	coin=coin+list(dy)
	dy=dy+1
	dy_cnt=dy_cnt+1
	i=i+1
	if dy=7 then dy=0
loop until ch=vbno