VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00ADA6A5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5280
   ClientLeft      =   540
   ClientTop       =   2235
   ClientWidth     =   5730
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5730
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   2400
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   400
      Left            =   1500
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2350
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Enabled         =   0   'False
      TickStyle       =   3
      TextPosition    =   1
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      Pattern         =   "*.wav*"
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00ADA6A5&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      ItemData        =   "Form1.frx":08CA
      Left            =   1200
      List            =   "Form1.frx":08CC
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00ADA6A5&
      Height          =   425
      Left            =   2550
      Picture         =   "Form1.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "ÎØæÉ ÃãÇãíø"
      Top             =   3360
      Width           =   425
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00ADA6A5&
      Height          =   425
      Left            =   3585
      Picture         =   "Form1.frx":0FF0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Stop Reader"
      Top             =   3360
      Width           =   425
   End
   Begin VB.CommandButton cmdPause 
      BackColor       =   &H00ADA6A5&
      Height          =   425
      Left            =   3075
      Picture         =   "Form1.frx":1332
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Pause Current Verse"
      Top             =   3360
      Width           =   425
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00ADA6A5&
      Height          =   425
      Left            =   2040
      Picture         =   "Form1.frx":1608
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "ØÇáÚ ÓõæúÑóÉ"
      Top             =   3360
      Width           =   425
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00ADA6A5&
      Height          =   425
      Left            =   1515
      Picture         =   "Form1.frx":1BCA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "ÎÜØæÉ ÎáÝí"
      Top             =   3360
      Width           =   425
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   1990
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   582
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      StepVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin VB.PictureBox Picture1 
      Height          =   7495
      Left            =   0
      Picture         =   "Form1.frx":20D4
      RightToLeft     =   -1  'True
      ScaleHeight     =   7440
      ScaleWidth      =   5715
      TabIndex        =   9
      Top             =   240
      Width           =   5775
      Begin MSComctlLib.Slider Slider2 
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Slide to Increase or decrease Speaker Volume"
         Top             =   4270
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   344
         _Version        =   393216
         MousePointer    =   4
         TickStyle       =   3
         TextPosition    =   1
      End
      Begin VB.CommandButton cmdShow 
         BackColor       =   &H00ADA6A5&
         Height          =   425
         Left            =   4800
         Picture         =   "Form1.frx":65EDA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Hide If Reading / Unhide Text Window If Stopped"
         Top             =   4320
         Width           =   425
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   4300
         Picture         =   "Form1.frx":6684C
         RightToLeft     =   -1  'True
         ScaleHeight     =   900
         ScaleWidth      =   720
         TabIndex        =   11
         Top             =   550
         Width           =   720
      End
      Begin VB.Label lbl5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lbl4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000016&
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lbl3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0E0FF&
         Height          =   225
         Left            =   1245
         TabIndex        =   14
         Top             =   2535
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label lbl2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1245
         TabIndex        =   13
         Top             =   1845
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÃÎÊÑ ÓæÑÉ ÇáÞÑÇä ÇáßÑíã ãä ÇáÞÇÆãÉ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1335
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   2865
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000A&
      FillColor       =   &H00ADA6A5&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1320
      Top             =   3240
      Width           =   2835
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hmixer  As Long
Dim VolCtrl As MIXERCONTROL
Dim varArr(), varI As Integer, varPos As Long, varPos1 As Long
Private Sub cmdNext_Click()
 
 If Me.Combo1.ListIndex < 113 Then
        
        Me.MMControl1.Command = "Stop"
        Me.Slider1.Value = 0
        Me.Slider1.Visible = False
        Me.lbl5.Caption = ""
        Me.lbl5.Visible = False
        
        Me.Combo1.ListIndex = Me.Combo1.ListIndex + 1
        
        Me.lbl2.Caption = ""
        Me.lbl3.Caption = ""
        Me.cmdKunoot.Enabled = False
        Me.cmdNaat.Enabled = False
  
        Call fncScroll
        
                 
 End If
 
End Sub
Private Sub cmdPause_Click()
  
  If Me.MMControl1.Mode = 524 Then Exit Sub
  
  If Me.MMControl1.Mode = 526 Then
    
    Me.MMControl1.Command = "Pause"
    varPos = Me.MMControl1.Position
    Me.cmdPause.ToolTipText = "Resume From Last Position"
    bln = True
    
  ElseIf Me.MMControl1.Mode = 529 Then
     
     Me.MMControl1.Command = "Resume"
     Me.MMControl1.Command = "Play"
     Me.MMControl1.Command = "From " & varPos
     Me.MMControl1.Command = "To " & Me.MMControl1.Length
     Me.cmdPause.ToolTipText = "Pause Current Verse"
     bln = False
     
  End If
  
  Me.cmdKunoot.Enabled = False
  Me.cmdNaat.Enabled = False
  
End Sub
Private Sub cmdPlay_Click()
   
   If Me.MMControl1.Mode = 526 Or Me.MMControl1.Mode = 529 Then Exit Sub
   
   If Combo1.ListIndex > 0 Then
     
     Call PlayBismillah(True)
     
   Else
     
     Me.Combo1.ListIndex = 0
     Call PlayBismillah(False)
     
   End If
   
   Me.Slider1.Value = 0
   Me.Slider1.Visible = True
   
End Sub
Private Sub cmdPrevious_Click()
   
  If Me.Combo1.ListIndex > 0 Then
                
        Me.MMControl1.Command = "Stop"
        Me.Slider1.Value = 0
        Me.Slider1.Visible = False
        Me.lbl5.Caption = ""
        Me.lbl5.Visible = False
        
        Me.Combo1.ListIndex = Me.Combo1.ListIndex - 1
        
        Me.lbl2.Caption = ""
        Me.lbl3.Caption = ""
    
        Call fncScroll
        
        Me.cmdKunoot.Enabled = False
        Me.cmdNaat.Enabled = False
  
  End If

   
End Sub
Private Sub cmdShow_Click()
    
    If frmAbout.WindowsMediaPlayer1.Visible = True Then Exit Sub
    
    If frmAbout.Visible = False Then
       
        frmAbout.Visible = True
        
        Me.Left = 500
        Me.Top = (Screen.Height - Me.Height) / 2
    
        frmAbout.WindowState = vbNormal
        frmAbout.Left = Me.Left + Me.Width
        frmAbout.Top = Me.Top + 400
        frmAbout.Height = 5100
        
    ElseIf frmAbout.Visible = True Then
        
        frmAbout.Visible = False
         
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2
        
    End If

End Sub
Private Sub cmdStop_Click()
 
 Me.Slider1.Value = 0
 Me.Slider1.Visible = False
 Me.MMControl1.Command = "Stop"
 
 frmAbout.RichTextBox1.TextRTF = ""
 frmAbout.RichTextBox1.Visible = False
 
 Me.cmdPlay.Enabled = True
 Me.Combo1.Locked = False
 
 Me.lbl2.Caption = ""
 Me.lbl3.Caption = ""
 Me.lbl5.Caption = ""
 Me.Slider2.Visible = False
 Me.lbl4.Visible = False
 varI = 0
 varPos = 0
 varPos1 = 0
 
 frmAbout.Timer1.Interval = 30
 Me.cmdKunoot.Enabled = True
 Me.cmdNaat.Enabled = True
  
 Call fncScroll
   
End Sub
Private Sub Combo1_Click()
     
    Me.Combo1.Refresh
  
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
 
 Dim I As Integer, varText As String, varchk As Boolean
 
 varchk = False
 
 If KeyAscii = 13 Then
    
    varText = Me.Combo1.Text
    
    For I = 1 To Me.Combo1.ListCount - 1
        
        If Left(Me.Combo1.List(I), Len(Me.Combo1)) = varText Then
           
           Me.Combo1.ListIndex = I
           Me.cmdPlay.SetFocus
           Me.cmdPlay.SetFocus
           varchk = True
           Exit For
            
               
        End If
    
    Next I
              
    If varchk = False Then
          
          Me.Combo1.Text = Me.Combo1.List(0)
        
    End If
    
          
 End If
 
End Sub

Private Sub Form_Load()
    
   Dim i1 As Integer, strItem As String, boK As Boolean, rc As Long
 
   Me.File1.Path = App.Path & "\AlBaset\"
   
   Do While i1 < 1
    
    strItem = File1.List(i1)
    strItem = Mid(strItem, 1, Len(strItem) - 4)
    Me.Combo1.AddItem strItem
    i1 = i1 + 1
      
   Loop
    
    frmPlayer.Show
    Me.Combo1.ListIndex = 0
    Me.cmdPlay.SetFocus
    
    Load frmAbout
    frmAbout.Show
    
    Me.Slider1.Visible = False
    Me.Slider2.Visible = False
    Me.lbl4.Visible = False
    Me.lbl5.Visible = False
    
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    
    If MMSYSERR_NOERROR <> rc Then
        
        Exit Sub
        
    End If
    
     boK = fGetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, VolCtrl)

     If boK Then
        
        With Slider2
            
            .Max = VolCtrl.lMaximum
            .Min = VolCtrl.lMinimum
            .Value = VolCtrl.lMaximum * 0.15
            .SmallChange = 1000
            .LargeChange = 1000
            
        End With
    
     End If
     
End Sub
Private Sub Form_Resize()
 
 If Me.WindowState = vbMinimized Then
    
    frmAbout.Visible = False
 
 ElseIf Me.WindowState = vbNormal Then
    
    Me.Left = 500
    Me.Top = (Screen.Height - Me.Height) / 2
    
    frmAbout.WindowState = vbNormal
    frmAbout.Left = Me.Left + Me.Width
    frmAbout.Top = Me.Top + 400
    frmAbout.Height = 5100
    
    frmAbout.Visible = True
    
 End If
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
  
   MMControl1.Command = "Close"
   MMControl1.Wait = True
   frmAbout.WindowsMediaPlayer1.Close
   Unload frmAbout
   End
   
End Sub
Function PlayBismillah(ByVal tf As Boolean)
    
    Dim chkSec As Double, varSec As Double
    
    Me.Combo1.Locked = True
    
    MMControl1.Command = "Close"
     
    frmAbout.PicScroll.Visible = False
    
    If frmAbout.WindowsMediaPlayer1.playState = wmppsStopped Then
       
       frmAbout.WindowsMediaPlayer1.Visible = False
    
    Else
       
       frmAbout.WindowsMediaPlayer1.Close
       frmAbout.WindowsMediaPlayer1.Visible = False
       
    End If
    
    frmAbout.RichTextBox1.Visible = True
    frmAbout.RichTextBox1.LoadFile App.Path & "\AlBaset\" & Left(Combo1.List(Combo1.ListIndex), 3) & ".rtf", rtfRTF
    frmAbout.Timer1.Enabled = True
     
    Me.lbl2.Caption = ""
    Me.lbl3.Caption = ""
    Me.Slider2.Visible = True
    Me.lbl4.Visible = True
    Me.lbl5.Visible = True
   
   If tf = True Then
     
     Me.MMControl1.Notify = False
     Me.MMControl1.Shareable = False
     Me.MMControl1.DeviceType = "WaveAudio"
   
     Me.MMControl1.FileName = App.Path & "\AlBaset\" & File1.List(114)
     Me.MMControl1.Wait = True
     Me.MMControl1.Command = "Open"
      
     Me.MMControl1.From = 0
     Me.MMControl1.To = Me.MMControl1.Length
     Me.MMControl1.Command = "Play"
   
     Do While Me.MMControl1.Position < Me.MMControl1.Length
        DoEvents
     Loop
 
   End If
    
   MMControl1.Command = "Close"
    
   Me.MMControl1.Notify = False
   Me.MMControl1.Shareable = False
   Me.MMControl1.DeviceType = "WaveAudio"
   Me.MMControl1.Wait = True
    
   Me.MMControl1.FileName = App.Path & "\AlBaset\" & Combo1.List(Combo1.ListIndex) & ".wav"
   Me.MMControl1.Wait = True
   Me.MMControl1.Command = "Open"
   
   Me.MMControl1.Command = "Play"
   Me.MMControl1.Wait = False
   Me.Slider1.Visible = True
   
   varI = 0
   varPos = 0
   varPos1 = 0
           
   Me.Timer1.Enabled = True
  
    Me.Timer1.Interval = 13000
  
    varSec = Int(Me.MMControl1.Length / 3600000)
    Me.lbl3.Caption = "[ Total Time << " & varSec & " h : "
         
    If Me.MMControl1.Length > 3600000 Then
     chkSec = Round(((Me.MMControl1.Length - (varSec * 3600000)) / 60000), 2)
    Else
     chkSec = Round((Me.MMControl1.Length / 60000), 2)
    End If
  
     Me.lbl3.Caption = Me.lbl3.Caption & Int(chkSec) & " m : "
     
     Me.lbl3.Caption = Me.lbl3.Caption & Int((chkSec - Int(chkSec)) * 60) & " Secs >> ]"
     Me.Slider1.Max = Me.MMControl1.Length

End Function
Private Sub MMControl1_StatusUpdate()
  
  Static varSec As Double
  Static chkSec As Double
  
    If Me.Slider1.Visible = False Then Exit Sub
    
    If Me.MMControl1.Position >= Me.MMControl1.Length Then
       
           Me.Slider1.Value = 0
           Me.lbl5.Caption = ""
           
           Call cmdStop_Click
           
    End If
       
       Me.Slider1.Value = Me.MMControl1.Position
       Me.lbl2.Visible = True
       Me.lbl3.Visible = True
             
       varSec = Int(Me.MMControl1.Position / 3600000)
         
       If Me.MMControl1.Position > 3600000 Then
        
        Me.lbl2.Caption = "[ Time Elapsed << " & varSec & " h : "
        chkSec = Round(((Me.MMControl1.Position - (varSec * 3600000)) / 60000), 2)
        
       Else
        
        chkSec = Round((Me.MMControl1.Position / 60000), 2)
        Me.lbl2.Caption = "[ Time Elapsed << "
        
       End If
  
        Me.lbl2.Caption = Me.lbl2.Caption & Int(chkSec) & " m : "
     
        Me.lbl2.Caption = Me.lbl2.Caption & Int((chkSec - Int(chkSec)) * 60) & " Secs >> ]"
                  
                  
End Sub
Private Sub Timer1_Timer()
 
 If Me.MMControl1.FileName = App.Path & "\AlBaset\" & File1.List(114) Then Exit Sub
 
 If Me.MMControl1.Mode = 526 Then
    
    varI = IIf(varI = 0, 1, varI)
    varPos = frmAbout.RichTextBox1.Find("(" & Trim(Str(varI)) & ")", varPos, , 10)

    If varPos = -1 Then Exit Sub

    If varI > 1 Then
       
       varPos1 = frmAbout.RichTextBox1.Find("(" & Trim(Str(varI - 1)) & ")", 0, , 10)
       frmAbout.RichTextBox1.SelStart = 0
       frmAbout.RichTextBox1.SelLength = varPos1
       frmAbout.RichTextBox1.SelColor = vbBlack

    End If
       
       frmAbout.RichTextBox1.SelStart = varPos1 + 3
       frmAbout.RichTextBox1.SelLength = varPos - varPos1
       frmAbout.RichTextBox1.SelColor = vbRed
 
       frmAbout.RichTextBox1.HideSelection = True
       varI = varI + 1
    
 End If
 
 
End Sub
Private Sub Slider2_Change()
    
    Static lVol As Long

    lVol = CLng(Slider2.Value)
    
    Call fSetVolumeControl(hmixer, VolCtrl, lVol)
    Me.lbl4.Caption = "Speaker Volume : " & Int(lVol * 100 / Slider2.Max) & "%"
   
End Sub
Private Sub Slider2_Scroll()
 Me.lbl4.Caption = "Speaker Volume : " & Int(Slider2.Value * 100 / Slider2.Max) & "%"
End Sub
Function IsLoaded(ByVal strFormName As String) As Boolean
 Dim I

   For I = 0 To Forms.Count - 1
      
      If Forms(I).Name = strFormName Then
         IsLoaded = True
      End If
      
   Next I


End Function
Private Sub fncScroll()
 
 frmAbout.RichTextBox1.TextRTF = ""
 frmAbout.RichTextBox1.Visible = False
       
 frmAbout.PicScroll.Visible = True
 
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
  
  frmAbout.PicScroll.FontSize = 10
  frmAbout.Timer1.Enabled = True
  
  RunMain

End Sub
 Private Function fSetVolumeControl(ByVal hmixer As Long, _
        mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
    Dim rc   As Long
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol  As MIXERCONTROLDETAILS_UNSIGNED

    With mxcd
        .item = 0
        .dwControlID = mxc.dwControlID
        .cbStruct = Len(mxcd)
        .cbDetails = Len(vol)
    End With

    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    vol.dwValue = volume

    Call CopyPtrFromStruct(mxcd.paDetails, vol, Len(vol))

    rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    Call GlobalFree(hmem)

    If MMSYSERR_NOERROR = rc Then
        fSetVolumeControl = True
    Else
        fSetVolumeControl = False
    End If

End Function



