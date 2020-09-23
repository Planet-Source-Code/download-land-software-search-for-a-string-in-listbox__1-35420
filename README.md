<div align="center">

## Search for a string in Listbox


</div>

### Description

Search for string in listbox
 
### More Info
 
This project needs a ListBox, named List1 and a TextBox, named Text1

the Listindex of the string that is found.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Download Land Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/download-land-software.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/download-land-software-search-for-a-string-in-listbox__1-35420/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Const LB_FINDSTRING = &H18F
```


### Source Code

```
Private Sub Form_Load()
  'KPD-Team 1998
  'URL: http://www.allapi.net/
  'E-Mail: KPDTeam@Allapi.net
  'Add some items to the listbox
  With List1
    .AddItem "Computer"
    .AddItem "Screen"
    .AddItem "Modem"
    .AddItem "Printer"
    .AddItem "Scanner"
    .AddItem "Sound Blaster"
    .AddItem "Keyboard"
    .AddItem "CD-Rom"
    .AddItem "Mouse"
  End With
End Sub
Private Sub Text1_Change()
  'Retrieve the item's listindex
  List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
End Sub
```

