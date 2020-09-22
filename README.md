<div align="center">

## Open/Close your CD\-ROM easily


</div>

### Description

Open and close your CDROM using the MCIsendstring API. I saw the program Heresy made, it's good but kind of complicated, newbies might not understand some stuff. I simply resolved the OPEN/CLOSE problem exploiting the Tag property of a command button. It's quite simple and reliable.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[$mTp ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mtp.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mtp-open-close-your-cd-rom-easily__1-7408/archive/master.zip)

### API Declarations

```
Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, ByVal _
lpstrReturnString As String, ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long
```


### Source Code

```
'create 2 command buttons, call the first one "Open" and the second one "Close"
'create a label
Private Sub Form_Load()
command1.tag = "open"
Private Sub Command1_Click()
If Command1.Tag = "open" Then
retvalue = mciSendString("set CDAudio door open", _
returnstring, 127, 0)
Command1.Tag = "closed"
Else
retvalue = mciSendString("set cdaudio door closed", returnstring, 127, 0)
Command1.Tag = "open"
End If
Label1.Caption = Command1.Tag 'place a label to check to tag property of the command button
```

