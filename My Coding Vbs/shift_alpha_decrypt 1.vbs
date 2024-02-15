p_txt=inputbox("Enter Plain Text")
enc_key=inputbox("Enter Key")

length=len(p_txt)
for i=1 to length
	txt=mid(p_txt,i,1)
	if txt=" " then msg=msg&" "
	if txt="a" then : shift = 0 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="b" then : shift = 1 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="c" then : shift = 2 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="d" then : shift = 3 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="e" then : shift = 4 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="f" then : shift = 5 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="g" then : shift = 6 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="h" then : shift = 7 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="i" then : shift = 8 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="j" then : shift = 9 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="k" then : shift = 10 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="l" then : shift = 11 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="m" then : shift = 12 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="n" then : shift = 13 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="o" then : shift = 14 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="p" then : shift = 15 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="q" then : shift = 16 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="r" then : shift = 17 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="s" then : shift = 18 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="t" then : shift = 19 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="u" then : shift = 20 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="v" then : shift = 21 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="w" then : shift = 22 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="x" then : shift = 23 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="y" then : shift = 24 - enc_key mod 26 : msg=msg & get_str(shift)
	if txt="z" then : shift = 25 - enc_key mod 26 : msg=msg & get_str(shift)
next

msgbox msg

function get_str(n)
	if n<0 then n=n+26
	set alpha_list=createobject("system.collections.arraylist")
	alpha_list.add "a"
	alpha_list.add "b"
	alpha_list.add "c"
	alpha_list.add "d"
	alpha_list.add "e"
	alpha_list.add "f"
	alpha_list.add "g"
	alpha_list.add "h"
	alpha_list.add "i"
	alpha_list.add "j"
	alpha_list.add "k"
	alpha_list.add "l"
	alpha_list.add "m"
	alpha_list.add "n"
	alpha_list.add "o"
	alpha_list.add "p"
	alpha_list.add "q"
	alpha_list.add "r"
	alpha_list.add "s"
	alpha_list.add "t"
	alpha_list.add "u"
	alpha_list.add "v"
	alpha_list.add "w"
	alpha_list.add "x"
	alpha_list.add "y"
	alpha_list.add "z"
	get_str=alpha_list(n)
end function