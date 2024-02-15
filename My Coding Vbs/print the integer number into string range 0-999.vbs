a=inputbox("enter")
if a=0 then msgbox "zero"
if a<=9 then
	msgbox one_nine(a)
elseif a>=10 and a<=20 then
	msgbox ten_twenty(a)
elseif a>=21 and a<=99 then
	msgbox twentyone_ninetynine(a)
elseif a>=100 and a<=999 then
	msgbox hundred_999(a)
end if
function one_nine(a)
	if a=1 then value="one"
	if a=2 then value="two"
	if a=3 then value="three"
	if a=4 then value="four"
	if a=5 then value="five"
	if a=6 then value="six"
	if a=7 then value="seven"
	if a=8 then value="eight"
	if a=9 then value="nine"
	one_nine=value
end function

function ten_twenty(a)
	if a=10 then value="ten"
	if a=11 then value="eleven"
	if a=12 then value="twelve"
	if a=13 then value="thirteen"
	if a=14 then value="fourteen"
	if a=15 then value="fifteen"
	if a=16 then value="sixteen"
	if a=17 then value="seventeen"
	if a=18 then value="eighteen"
	if a=19 then value="nineteen"
	if a=20 then value="twenty"
	ten_twenty=value
end function

function data_20_90s(a)
	digit=mid(a,1,1)
	digit2=mid(a,2,1)
	if digit=2 then value="twenty "
	if digit=3 then value="thirty "
	if digit=4 then value="fourty "
	if digit=5 then value="fifty "
	if digit=6 then value="sixty "
	if digit=7 then value="seventy "
	if digit=8 then value="eighty "
	if digit=9 then value="ninety "
	data_20_90s=value
end function

function twentyone_ninetynine(a)
	digit=mid(a,1,1)
	digit2=mid(a,2,1)
	if digit=2 then value=data_20_90s(digit)&one_nine(digit2)
	if digit=3 then value=data_20_90s(digit)&one_nine(digit2)
	if digit=4 then value=data_20_90s(digit)&one_nine(digit2)
	if digit=5 then value=data_20_90s(digit)&one_nine(digit2)
	if digit=6 then value=data_20_90s(digit)&one_nine(digit2)
	if digit=7 then value=data_20_90s(digit)&one_nine(digit2)
	if digit=8 then value=data_20_90s(digit)&one_nine(digit2)
	if digit=9 then value=data_20_90s(digit)&one_nine(digit2)
	twentyone_ninetynine=value
end function

function hundred_999(a)
	digit=mid(a,1,1)
	digit2=mid(a,2,1)
	digit3=mid(a,3,1)
	digit2v=mid(a,2,3)
	if digit=1 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=2 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=3 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=4 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=5 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=6 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=7 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=8 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=9 and digit2v>19 then value=one_nine(digit)&" hundred "&data_20_90s(digit2)&one_nine(digit3)
	if digit=1 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=2 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=3 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=4 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=5 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=6 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=7 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=8 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=9 and digit2v<10 then value=one_nine(digit)&" hundred "&one_nine(digit3)
	if digit=1 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	if digit=2 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	if digit=3 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	if digit=4 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	if digit=5 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	if digit=6 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	if digit=7 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	if digit=8 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	if digit=9 and digit2v>=10 and digit2v<=19 then value=one_nine(digit)&" hundred "&ten_twenty(digit2v)
	hundred_999=value
end function