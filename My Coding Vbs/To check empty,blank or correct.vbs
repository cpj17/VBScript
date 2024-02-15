do
a=inputbox("enter")
if a="cpj" then
msgbox "correct"
exit do
elseif isempty(a) then
wscript.quit
elseif a="" then
msgbox "it is blank"
else
msgbox"pls enetr correct"
end if
loop