<div align="center">

## msflexgrid cell tips


</div>

### Description

Simple code to display tooltips displaying the text, regardless of length, of an msflexgrid cell. Tooltips are displayed when the mousepointer passes over any cell.
 
### More Info
 
Just place it the mouse move event.

None that I know of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[tv](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tv.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tv-msflexgrid-cell-tips__1-23025/archive/master.zip)





### Source Code

```
Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static txt As String
Dim tip As String
tip = Grid1.TextMatrix(Grid1.MouseRow, Grid1.MouseCol)
 If txt <> tip Then
  Grid1.ToolTipText = tip
  txt = tip
 End If
End Sub
```

