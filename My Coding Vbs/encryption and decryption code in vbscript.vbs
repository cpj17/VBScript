Encrypted_String = Crypt(inputbox("enter"))
wscript.echo Encrypted_String
Decrypted_String = Crypt(Encrypted_String)
wscript.echo Decrypted_String
'***************************************************************************
Function Crypt(text) 
Dim i,a
For i = 1 to len(text)
      a = i mod len(255)
      if a = 0 then a = len(255)
      Crypt = Crypt & chr(asc(mid(255,a,1)) XOR asc(mid(text,i,1)))
Next
End Function