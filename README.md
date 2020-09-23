<div align="center">

## Open a text file FAST\!


</div>

### Description

Opens a text file much fast than a "Do While Not EOF(filenumber)" loop. Makes

file load times almost non-existant.
 
### More Info
 
The filename of the file to open.

The filename that you pass to the function must exist, otherwise an error will occur.

The text contained in the file.

I have yet to run into any. If you do happen to find one leave a comment and

I will fix it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Timothy Pew](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/timothy-pew.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/timothy-pew-open-a-text-file-fast__1-1732/archive/master.zip)





### Source Code

```
Public Function OpenFile(ByVal file As String) As String
 Dim i As Integer
 i = FreeFile
 Open file For Input As #i
 OpenFile = Input(LOF(i), i)
 Close #i
End Function
```

