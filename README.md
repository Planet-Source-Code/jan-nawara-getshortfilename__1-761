<div align="center">

## GetShortFileName


</div>

### Description

A short pathname of the passed string containing a long pathname.

'For example it turns "C:\Windows\MY Long Path Name\My Long Name File.txt" into "c:\windows\mylong~1\mylong~1.txt" (The actual resulting pathname is determined by the short names that windows assigns to all files and directories).

'This is useful when you need to create a fail proof pathname (assuming the file exists and is accesible).
 
### More Info
 
Requires that a pathname be passed.

A short DOS 8.3 format pathname.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jan Nawara](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jan-nawara.md)
**Level**          |Beginner
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jan-nawara-getshortfilename__1-761/archive/master.zip)

### API Declarations

```
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
```


### Source Code

```
Public Function Short_Name(Long_Path As String) As String
'Returns short pathname of the passed long pathname
Dim Short_Path As String
Dim PathLength As Long
Short_Path = Space(250)
PathLength = GetShortPathName(Long_Path, Short_Path, Len(Short_Path))
If PathLength Then
 Short_Name = Left$(Short_Path, PathLength)
End If
End Function
```

