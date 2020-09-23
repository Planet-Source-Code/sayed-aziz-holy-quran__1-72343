Attribute VB_Name = "mDeclaration"
Option Explicit

Public bln As Boolean, varPos As Long, varFileTxt As String
Public ScrollText As String
Public rt As Long
Public DrawingRect As RECT
Public UpperX As Long, UpperY As Long
Public RectHeight As Long
Public Sub RunMain()

Const IntervalTime As Long = 20
 
frmAbout.Refresh
 
rt = DrawText(frmAbout.PicScroll.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then
    
Else
   
    DrawingRect.Top = frmAbout.PicScroll.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = frmAbout.PicScroll.ScaleWidth
   
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + frmAbout.PicScroll.ScaleHeight
    
    
    
End If

End Sub

