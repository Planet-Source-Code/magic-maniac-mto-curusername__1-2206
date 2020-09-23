<div align="center">

## CurUserName


</div>

### Description

Get the current username from Windows
 
### More Info
 
Dim sStr1$

sStr1 = CurUserName()

( The current username from Windows )

sStr1 = "MAGiC MANiAC^mTo"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MAGiC MANiAC^mTo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/magic-maniac-mto.md)
**Level**          |Advanced
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/magic-maniac-mto-curusername__1-2206/archive/master.zip)





### Source Code

```
' Get the current username from Windows
' Coded By MAGiC MANiAC^mTo
' More Examples At: http://home.kabelfoon.nl/~mto/
'
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _ (ByVal lpBuffer As String, nSize As Long) As Long
Function CurUserName$()
 Dim sTmp1$
 sTmp1 = Space$(512)
 GetUserName sTmp1, Len(sTmp1)
 CurUserName = Trim$(sTmp1)
End Function
```

