<div align="center">

## How to use the common dialog control


</div>

### Description

Use of 3 types of common dialog boxes:1: choose printer, 2: choose font, 3: choose color. http://137.56.41.168:2080/VisualBasicSource/vb4usecommondialog.txt
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-how-to-use-the-common-dialog-control__1-671/archive/master.zip)





### Source Code

```
'1: choose printer
Public Sub ChoosePrinter
  Const ErrCancel = 32755
  CommonDialog1.CancelError = True
On Error GoTo errorPrinter
  CommonDialog1.Flags = 64
  'see the Help on Flags Properties (Print Dialog)
  CommonDialog1.ShowPrinter
  CommonDialog1.PrinterDefault = False
  Exit Sub
errorPrinter:
  If Err = ErrCancel Then Exit Sub Else Resume
End Sub
'2: choose font
Global vScreenFont, vScreenFontSize
Public Sub ChooseFont()
  CommonDialog1.Flags = cdlCFScreenFonts
  'see the Help on Flags Properties (Font Dialog)
  CommonDialog1.ShowFont
  vScreenFont = CommonDialog1.FontName
  vScreenFontSize = CommonDialog1.FontSize
  Call ChangeFont(Form1)
End Sub
Public Sub ChangeFont(X As Form)
  Dim Control
  For Each Control In X.Controls
    If TypeOf Control Is Label Or _
      TypeOf Control Is TextBox Or _
      TypeOf Control Is CommandButton Or _
      TypeOf Control Is ComboBox Or _
      TypeOf Control Is ListBox Or _
      TypeOf Control Is CheckBox Then
        Control.Font = vScreenFont
        Control.FontSize = vScreenFontSize
    End If
  Next Control
End Sub
'3: choose color
Global vColor
Public Sub ChooseColor
  CommonDialog1.Flags = &H1& Or &H4&
  'see the Help on Flags Properties (Color Dialog)
  CommonDialog1.ShowColor
  vColor = CommonDialog1.Color
'  if you want to convert the color to hex use
'  MsgBox Convert2Hex(vColor)
'  if you want to repaint youre background use
'  Call ChangeColor(X as Form)
End Sub
Public Sub ChangeColor(X As Form)
  Dim Control
  X.BackColor = vColor
  For Each Control In X.Controls
    If TypeOf Control Is Label Or _
      TypeOf Control Is TextBox Or _
      TypeOf Control Is CommandButton Or _
      TypeOf Control Is ComboBox Or _
      TypeOf Control Is ListBox Or _
      TypeOf Control Is CheckBox Then
        Control.BackColor = vColor
    End If
  Next Control
End Sub
Public Function Convert2Hex(color) as String
	Dim RedValue, GreenValue, BlueValue
    RedValue = (color And &HFF&)
    GreenValue = (color And &HFF00&) \ 256
    BlueValue = (color And &HFF0000) \ 65536
    Convert2Hex = Format(Hex(RedValue) & Hex(GreenValue) & Hex(BlueValue), "000000")
End Function
```

