Attribute VB_Name = "Progress"
Public Function Progress(PictureProgress As PictureBox, MaxLength As Long, Value As Long)
Dim i As Long
    i = (100 * Value) / MaxLength
    
PictureProgress.ForeColor = RGB(50, 50, 150)
FillIt PictureProgress, i
    
End Function
Public Sub FillIt(PicBox As Control, percent As Long)
Dim PercSh$
    
If Not PicBox.AutoRedraw Then
    PicBox.AutoRedraw = -1
End If
    
PicBox.Cls
PicBox.ScaleWidth = 100
PicBox.DrawMode = 10
PercSh$ = Format$(percent, "###") + "%"
PicBox.CurrentX = 50 - PicBox.TextWidth(PercSh$) / 2
PicBox.CurrentY = (PicBox.ScaleHeight - PicBox.TextHeight(PercSh$)) / 2
PicBox.Print PercSh$
PicBox.Line (0, 0)-(percent, PicBox.ScaleHeight), , BF
PicBox.Refresh

End Sub



