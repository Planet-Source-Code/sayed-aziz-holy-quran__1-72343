VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5970
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   5970
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   1680
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
   Exit Sub
   If IsLoaded("frmSplash") Then
    
    Unload Me
    Load frmPlayer
    frmPlayer.Show
   
   End If
   
End Sub
Function IsLoaded(ByVal strFormName As String) As Boolean
 Dim I   ' Declare variable.
   ' Refill list (in case an instance was added or removed).

   For I = 0 To Forms.Count - 1
      
      If Forms(I).Name = strFormName Then
         IsLoaded = True
      End If
      
   Next I


End Function
Private Sub Form_Load()
 Me.Timer1.Interval = 500
 SetIcon Me.hwnd, "101", True
End Sub
Private Sub Timer1_Timer()
 Static I
 I = I + 1
 
 If I > 5 Then
  Unload Me
    Load frmPlayer
    frmPlayer.Show
 End If
 
End Sub
