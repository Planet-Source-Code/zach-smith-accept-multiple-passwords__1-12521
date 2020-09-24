<div align="center">

## Accept multiple passwords\.


</div>

### Description

Accept multiple passwords, yet still deny access for wrong password entries.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[zach smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/zach-smith.md)
**Level**          |Beginner
**User Rating**    |3.0 (9 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/zach-smith-accept-multiple-passwords__1-12521/archive/master.zip)





### Source Code

```
Dim Pw1 As String
Dim Pw2 As String
Pw1 = "password1"
Pw2 = "password2"
If Text1 = Pw1 Then
MsgBox "you entered the correct password"
Else
If Text1 = Pw2 Then
MsgBox "you entered the correct password"
Else
MsgBox "wrong password, please try again"
End If
End If
```

