VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   195
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   120
      Top             =   4440
   End
   Begin VB.PictureBox PicScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4875
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      RightToLeft     =   -1  'True
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   4875
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   5295
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   -1  'True
         enabled         =   -1  'True
         enableContextMenu=   0   'False
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   9340
         _cy             =   8599
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5055
      Left            =   10
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5500
      _ExtentX        =   9710
      _ExtentY        =   8916
      _Version        =   393217
      BackColor       =   14737632
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAbout.frx":61412
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 On Error Resume Next
  
  Me.RichTextBox1.Visible = False
  Me.WindowsMediaPlayer1.URL = App.Path & "\MB.wmv"
  Me.WindowsMediaPlayer1.enableContextMenu = False
  
  
End Sub
Private Sub picScroll_Click()
 Timer1.Enabled = Not Timer1.Enabled
End Sub
Private Sub Timer1_Timer()
  
        PicScroll.Cls
        
        DrawText PicScroll.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
       
        If DrawingRect.Top < -(RectHeight) Then
            DrawingRect.Top = PicScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + PicScroll.ScaleHeight
        End If
        
        PicScroll.AutoRedraw = True
    
   If Me.WindowsMediaPlayer1.Visible = True Then Me.Refresh
   
End Sub
Private Sub WindowsMediaPlayer1_StatusChange()
 
 If Me.WindowsMediaPlayer1.playState = wmppsStopped Then
    
        If Me.WindowsMediaPlayer1.URL = App.Path & "\MB.wmv" Then
      
            Me.WindowsMediaPlayer1.Close
            Me.WindowsMediaPlayer1.URL = App.Path & "\Naat.wma"
            DoEvents
            
        Else
        
            Me.WindowsMediaPlayer1.Close
            Me.WindowsMediaPlayer1.Visible = False
            Me.PicScroll.Visible = True
    
            ScrollText = "Quran Reader " & vbCrLf & vbCrLf & _
                "Version 1.1.0" & vbCrLf & vbCrLf & _
                "Developed By:" & vbCrLf & _
                "<< Sayed Aziz Ahmed >>" & vbCrLf & _
                " Riyadh, Saudi Arabia " & vbCrLf & _
                " Cell : 0551468967 " & vbCrLf & vbCrLf & _
                " In July 2008 " & vbCrLf & vbCrLf & _
                " When you Select Quran Verse " & vbCrLf & _
                " From Drop Down Menu," & vbCrLf & _
                " Quran Text will appear here " & vbCrLf & vbCrLf & _
                " If you have any better ideas," & vbCrLf & _
                " comments, suggestions etc.," & vbCrLf & _
                " You can email me." & vbCrLf & vbCrLf & _
                " E-mail : aziz_abroad@yahoo.com" & vbCrLf & _
                " Web Site :" & vbCrLf & _
                " www.geocities.com/aziz_abroad" & vbCrLf
                                  
            Me.Timer1.Interval = 20
            Timer1.Enabled = True
    
            RunMain
  
        End If
    
    End If
 
End Sub
Private Sub WmStopped(ByVal fname As String)

 Select Case fname

    Case App.Path & "\MB.wmv"
      
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB1.mp3"
        DoEvents
      
    Case App.Path & "\MB1.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB2.mp3"
        DoEvents
        
    Case App.Path & "\MB2.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB3.mp3"
        DoEvents
        
    Case App.Path & "\MB3.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB4.mp3"
        DoEvents
        
    Case App.Path & "\MB4.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB5.mp3"
        DoEvents
        
    Case App.Path & "\MB5.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB6.mp3"
        DoEvents
        
    Case App.Path & "\MB6.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB7.mp3"
        DoEvents
        
    Case App.Path & "\MB7.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB8.mp3"
        DoEvents
        
    Case App.Path & "\MB8.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB9.mp3"
        DoEvents
        
    Case App.Path & "\MB9.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB10.mp3"
        DoEvents
        
    Case App.Path & "\MB10.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB11.mp3"
        DoEvents
    
    Case App.Path & "\MB11.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB12.mp3"
        DoEvents
        
    Case App.Path & "\MB12.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB13.mp3"
        DoEvents
    
    Case App.Path & "\MB13.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB14.mp3"
        DoEvents
        
    Case App.Path & "\MB14.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB15.mp3"
        Me.WindowsMediaPlayer1.SetFocus
        DoEvents
        
    Case App.Path & "\MB15.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB16.mp3"
        DoEvents
    
    Case App.Path & "\MB16.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB17.mp3"
        DoEvents
    
    Case App.Path & "\MB17.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB18.mp3"
        DoEvents
        
    Case App.Path & "\MB18.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB19.mp3"
        DoEvents
    
    Case App.Path & "\MB19.mp3"
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.URL = App.Path & "\MB20.mp3"
        DoEvents
        
    Case Else
        
        Me.WindowsMediaPlayer1.Close
        Me.WindowsMediaPlayer1.Visible = False
        Me.PicScroll.Visible = True
    
        ScrollText = "Quran Reader " & vbCrLf & vbCrLf & _
                "Version 1.1.0" & vbCrLf & vbCrLf & _
                "Developed By:" & vbCrLf & _
                "<< Sayed Aziz Ahmed >>" & vbCrLf & _
                " Riyadh, Saudi Arabia " & vbCrLf & _
                " Cell : 0551468967 " & vbCrLf & vbCrLf & _
                " In July 2008 " & vbCrLf & vbCrLf & _
                " When you Select Quran Verse " & vbCrLf & _
                " From Drop Down Menu," & vbCrLf & _
                " Quran Text will appear here " & vbCrLf & vbCrLf & _
                " If you have any better ideas," & vbCrLf & _
                " comments, suggestions etc.," & vbCrLf & _
                " You can email me." & vbCrLf & vbCrLf & _
                " E-mail : aziz_abroad@yahoo.com" & vbCrLf & _
                " Web Site :" & vbCrLf & _
                " www.geocities.com/aziz_abroad" & vbCrLf
                                  
        Me.Timer1.Interval = 20
        Timer1.Enabled = True
    
        RunMain
  
    End Select
    
End Sub
