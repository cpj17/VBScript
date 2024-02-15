listPath = "C:\Users\GSIAD-031\Desktop\LocalDeployPaths.vbs"

On Error Resume Next
set obj = CreateObject("Wscript.Shell")
set fso = CreateObject("Scripting.FileSystemObject")
'set vbF = fso.OpenTextFile("C:\Users\"& obj.ExpandEnvironmentStrings("%USERNAME%") &"\Desktop\"& listPath)
set vbF = fso.OpenTextFile(listPath)
Execute vbF.readall()
vbf.close()

i= 1
For Each key in KeyList
	keys = keys & i &". "& key & vbLf
	i = i + 1
Next

keys = keys &vbLf&"99 To Exit."

Main
Sub Main
	opt = inputbox("Enter Your Option" &vbLf&vbLf& keys, "Made By CPJ", DefaultValue)

	If IsNumeric(opt) Then
		If opt <> 99 Then
			obj.run valueList(opt - 1)
		Else
			Wscript.Quit
		End If
	Else
		MsgBox "Please Chose Valid Choice"
		Main()
	end if
End Sub