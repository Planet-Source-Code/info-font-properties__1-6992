<div align="center">

## ^ Font Properties ^


</div>

### Description

This little coding lets the user see how many fonts they have and what they can do with it..this is kind of old..but i was bored.
 
### More Info
 
3 - CommandButtons (Command1, Command2, Command3)

1 - Listbox (List1)

1 - Label (Label1)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[iNfO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/info.md)
**Level**          |Beginner
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/info-font-properties__1-6992/archive/master.zip)





### Source Code

```
'- Made By: iNfO
'- Project: ^ Font Properties ^
'- Items Needed:
'-       3 - CommandButtons (Command1, Command2, Command3)
'-       1 - Listbox (List1)
'-       1 - Label (Label1)
Private Sub Command1_Click()
'- Declares the variables
Dim NUM As Single
Dim x As Single
'- gets the numbers of fonts you have
NUM = Screen.FontCount
'- Set the listbox properties
'- Set List1, Sorted = True
'- Goes from 1 to number of fonts
For x = 1 To NUM
  List1.AddItem Screen.Fonts(x)
Next x
'- for some reason there will be a blank itme
'- this removes it
List1.RemoveItem (0)
'- Displays the number of fonts
Label2.Caption = List1.ListCount
End Sub
Private Sub Command2_Click()
'- Makes sure that there are fonts to choose from
If List1.ListCount <> 0 Then
  '- this makes the fonts watever you select from
  '- the listbox
  Label1.Font = List1.Text
Else
  MsgBox "you have to choose the fonts first"
End If
End Sub
Private Sub Command3_Click()
'- Makes sure that there are fonts to choose from
If List1.ListCount <> 0 Then
  '- Declares the variables
  Dim Size As Single
  '- lets it inputbox get the font size
  '- Makes it a value
  Size = Val(InputBox("Enter the font size"))
  Label1.FontSize = Val(Size)
Else
  MsgBox "you have to choose the fonts first"
End If
End Sub
Private Sub Form_Load()
'- Sets the captions of the buttons
Command1.Caption = "Get Fonts"
Command2.Caption = "Apply Fonts"
Command3.Caption = "Get Fonts Size"
End Sub
```

