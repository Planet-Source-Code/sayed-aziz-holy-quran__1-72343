Attribute VB_Name = "mVisual"
Option Explicit

Public Const DT_CALCRECT As Long = &H400
Public Const DT_CENTER As Long = &H1
Public Const DT_WORDBREAK As Long = &H10
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50

Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" _
        (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, _
         ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuload As Long) As Long
         
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4
Public Sub SetIcon(ByVal hwnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
    Dim lhwndTop As Long
    Dim lhwnd As Long
    Dim cx As Long
    Dim cy As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long
 
    If bSetAsAppIcon Then
    
        lhwnd = hwnd
        lhwndTop = lhwnd
    
        Do While Not (lhwnd = 0)
            lhwnd = GetWindow(lhwnd, GW_OWNER)
        
            If Not (lhwnd = 0) Then
                lhwndTop = lhwnd
            End If
        
        Loop
        
     End If
     
     cx = GetSystemMetrics(SM_CXICON)
     cy = GetSystemMetrics(SM_CYICON)
     
     hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
  
     If bSetAsAppIcon Then
        SendMessageLong lhwndTop, WM_SETICON, ICON_BIG, hIconLarge
     End If
        
        SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
        
        cx = GetSystemMetrics(SM_CXSMICON)
        cy = GetSystemMetrics(SM_CYSMICON)
        hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
      
     If bSetAsAppIcon Then
        
        SendMessageLong lhwndTop, WM_SETICON, ICON_SMALL, hIconSmall
     
     End If
        
        SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
     
End Sub


