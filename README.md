<div align="center">

## Get Shortcut's Target


</div>

### Description

After looking all over on PSC i was unable to find Short, Simple and Clean code to get the target path of a window's shortcut (.lnk) file, so here is an easier way.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael L\. Canejo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-l-canejo.md)
**Level**          |Beginner
**User Rating**    |4.2 (46 globes from 11 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-l-canejo-get-shortcut-s-target__1-44596/archive/master.zip)





### Source Code

```

Public Function GetTarget(strPath As String) As String
  'Gets target path from a shortcut file
On Error GoTo Error_Loading
  Dim wshShell As Object
  Dim wshLink As Object
  Set wshShell = CreateObject("WScript.Shell")
  Set wshLink = wshShell.CreateShortcut(strPath)
  GetTarget = wshLink.TargetPath
  Set wshLink = Nothing
  Set wshShell = Nothing
  Exit Function
Error_Loading:
  GetTarget = "Error occured."
End Function
```

