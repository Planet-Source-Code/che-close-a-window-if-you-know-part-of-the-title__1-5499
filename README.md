<div align="center">

## Close a window\(if you know part of the title\)


</div>

### Description

This code closes a window, if you know part of it's title. It uses some AOL API's, hehe(sendmessagebystring). I used it for closing Netscape windows because the "Window Class Name" always changed. Not sure if this code has any use though...
 
### More Info
 
1. The title, or part of the title.

2. A form named "Form1", or you could change the code a little


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Che](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/che.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/che-close-a-window-if-you-know-part-of-the-title__1-5499/archive/master.zip)

### API Declarations

```
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Const WM_CLOSE = &H10
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
```


### Source Code

```
Function FindWindowByTitle(Title As String)
Dim a, b, Caption
 a = getwindow(Form1.hWnd, GW_OWNER)
 Caption = GetCaption(a)
 If InStr(1, LCase(Caption), LCase(Title)) <> 0 Then
  FindWindowByTitle = b
  Exit Function
 End If
 b = a
 Do While b <> 0: DoEvents
  b = getwindow(b, GW_HWNDNEXT)
  Caption = GetCaption(b)
  If InStr(1, LCase(Caption), LCase(Title)) <> 0 Then
   FindWindowByTitle = b
   Exit Do
   Exit Function
  End If
 Loop
End Function
Function GetCaption(hWnd)
 dim hwndLength%, hwndTitle$, a%
 hwndLength% = GetWindowTextLength(hWnd)
 hwndTitle$ = String$(hwndLength%, 0)
 a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))
 GetCaption = hwndTitle$
End Function
Sub KillWin(Title As String)
Dim a, hWnd
 hWnd = FindWindowByTitle(Title)
 a = sendmessagebystring(hWnd, WM_CLOSE, 0, 0)
End Sub
Use KillWin to close the window.
```

