<div align="center">

## Semi transparent form 2000/XP


</div>

### Description

Your standard semi transparent form in Win 2000/XP without the flickering. Comes with fade-in effect.
 
### More Info
 
Must be running Win 2000 equivalent or better

1 form needed - copy everything onto it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kjhel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kjhel.md)
**Level**          |Beginner
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kjhel-semi-transparent-form-2000-xp__1-67024/archive/master.zip)

### API Declarations

```
Option Explicit
'
' Constants
'
Private Const LWA_ALPHA  As Long = &amp;H2
Private Const WS_EX_LAYERED As Long = &amp;H80000
Private Const GWL_EXSTYLE  As Long = -20
Private Const SW_SHOW   As Long = 5
Private Const RDW_UPDATENOW As Long = &amp;H100
'
' Declarations
'
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, _
             lprcUpdate As Any, _
             ByVal hrgnUpdate As Long, _
             ByVal fuRedraw As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                   ByVal nIndex As Long, _
                   ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                   ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
             ByVal nCmdShow As Long) As Long
```


### Source Code

```
Private Sub Form_Load()
 Trans Me.hwnd, 1, 85
End Sub
Private Sub Trans(lngHwnd As Long, _
     Optional ByVal Speed As Byte = 1, _
     Optional ByVal OpaquePercent As Byte = 85)
Dim Cnt As Long
 On Error Resume Next
 'Layered window
 SetWindowLong lngHwnd, GWL_EXSTYLE, GetWindowLong(lngHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
 SetLayeredWindowAttributes lngHwnd, 0, 0, LWA_ALPHA
 'Smoothness
 ShowWindow lngHwnd, SW_SHOW
 RedrawWindow lngHwnd, ByVal 0&, ByVal 0&, RDW_UPDATENOW
 'Fade-in effect
 'OpaquePercent range = 0 to 100 [Default = 85]
 'Speed range = 0 to about 5 [Default = 1] (visible difference not much for higher values)
 For Cnt = 0 To OpaquePercent Step Speed
  SetLayeredWindowAttributes lngHwnd, 0, (Cnt / 100) * 255, LWA_ALPHA
 Next Cnt
End Sub
```

