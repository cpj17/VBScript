set obj = CreateObject("wscript.shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Function pd(n, totalDigits) 
	if totalDigits > len(n) then 
		pd = String(totalDigits-len(n),"0") & n 
	else 
		pd = n 
	end if 
End Function

curDate = Pd(DAY(date()),2) & Pd(Month(date()),2) & YEAR(Date())

if objFSO.folderexists("\\192.168.0.57\Gurusoft\Server\Logs\TableLogs\G10WMSR1V2MOBILE\" & curDate & "\SGH\ADMIN") then
	cmd = "\\192.168.0.57\Gurusoft\Server\Logs\TableLogs\G10WMSR1V2MOBILE\" & curDate & "\SGH\ADMIN"
else
	cmd = "\\192.168.0.57\Gurusoft\Server\Logs\TableLogs\G10WMSR1V2MOBILE\"
end if

obj.run cmd