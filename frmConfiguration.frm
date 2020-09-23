VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dilbert's On-Line Screen Saver Configuration"
   ClientHeight    =   2460
   ClientLeft      =   3540
   ClientTop       =   2835
   ClientWidth     =   9330
   Icon            =   "frmConfiguration.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   9330
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton butSelectColour 
      Caption         =   "&Select..."
      Height          =   390
      Left            =   7995
      TabIndex        =   8
      Top             =   1455
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox picBackgroundSample 
      Enabled         =   0   'False
      Height          =   1380
      Index           =   0
      Left            =   6195
      Picture         =   "frmConfiguration.frx":08CA
      ScaleHeight     =   1320
      ScaleWidth      =   2955
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5355
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrCheckParent 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5115
      Top             =   645
   End
   Begin VB.Timer tmrCaptionUpdate 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5535
      Top             =   645
   End
   Begin VB.CheckBox chkHideCursor 
      Caption         =   "The cursor will be &hidden."
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   1215
      Width           =   2415
   End
   Begin VB.ComboBox cmbDesktopEffect 
      Height          =   315
      ItemData        =   "frmConfiguration.frx":2DE6
      Left            =   7260
      List            =   "frmConfiguration.frx":2DE8
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   150
      Width           =   1965
   End
   Begin VB.CheckBox chkIsManyCartoonsDIsplayed 
      Caption         =   "&Old cartoons will be left on the screen..."
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   615
      Width           =   3180
   End
   Begin VB.CommandButton butOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   7905
      TabIndex        =   15
      Top             =   2040
      Width           =   1245
   End
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6555
      TabIndex        =   14
      Top             =   2040
      Width           =   1245
   End
   Begin VB.ComboBox cmbCartoonInterval 
      Height          =   315
      ItemData        =   "frmConfiguration.frx":2DEA
      Left            =   3885
      List            =   "frmConfiguration.frx":2DEC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   165
      Width           =   1515
   End
   Begin VB.CheckBox chkDisplayRefersh 
      Caption         =   "but the screen will be &cleared every 10 minutes."
      Height          =   240
      Left            =   495
      TabIndex        =   6
      Top             =   870
      Width           =   3690
   End
   Begin VB.PictureBox picBackgroundSample 
      Enabled         =   0   'False
      Height          =   1380
      Index           =   2
      Left            =   6195
      Picture         =   "frmConfiguration.frx":2DEE
      ScaleHeight     =   1320
      ScaleWidth      =   2955
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picBackgroundSample 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   1380
      Index           =   1
      Left            =   6195
      ScaleHeight     =   1320
      ScaleWidth      =   2955
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label labWeblink 
      Caption         =   "Dilbert Homepage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   735
      TabIndex        =   13
      Top             =   2085
      Width           =   1305
   End
   Begin VB.Label Label3 
      Caption         =   "Visit the                                for more related fun!"
      Height          =   285
      Left            =   90
      TabIndex        =   12
      Top             =   2085
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Dilbert"
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   225
      Width           =   570
   End
   Begin VB.Label Label6 
      Caption         =   "and the &backdrop will be"
      Height          =   255
      Left            =   5445
      TabIndex        =   3
      Top             =   210
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   570
      Picture         =   "frmConfiguration.frx":23EB4
      ToolTipText     =   "Hey! Get that cursor off my face, man!"
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "will show you a new cartoon strip &every"
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Top             =   225
      Width           =   3180
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub butCancel_Click()
   Unload Me
End Sub

Private Sub butOk_Click()
   Dim i1%, s1$, s2$
   On Error Resume Next
   
   Select Case cmbCartoonInterval.ListIndex
      Case 0
         s1 = "10SEC"
      Case 1
         s1 = "20SEC"
      Case 2
         s1 = "30SEC"
      Case 3
         s1 = "1MIN"
      Case 4
         s1 = "2MIN"
      Case 5
         s1 = "10MIN"
      Case Else
         s1 = "20SEC"
   End Select
      
   SaveSetting App.EXEName, "Configuration", "DisplayDuration", s1
   
   Select Case cmbDesktopEffect.ListIndex
      Case 0
         s1 = "DESK"
      Case 1
         s1 = "BLK"
      Case 2
         s1 = "DIL"
      Case Else
         s1 = "DESK"
   End Select
   
   SaveSetting App.EXEName, "Configuration", "BackdropEffect", s1
      
   If chkIsManyCartoonsDIsplayed.Value = 1 Then
      SaveSetting App.EXEName, "Configuration", "MultipleCartoons", "YES"
   Else
      SaveSetting App.EXEName, "Configuration", "MultipleCartoons", "NO"
   End If
   
   If chkDisplayRefersh.Value = 1 Then
      SaveSetting App.EXEName, "Configuration", "MultipleRefresh", "YES"
   Else
      SaveSetting App.EXEName, "Configuration", "MultipleRefresh", "NO"
   End If
   
   If chkHideCursor.Value = 1 Then
      SaveSetting App.EXEName, "Configuration", "HideCursor", "YES"
   Else
      SaveSetting App.EXEName, "Configuration", "HideCursor", "NO"
   End If
      
   SaveSetting App.EXEName, "Configuration", "BackgroundColour", Hex$(picBackgroundSample(1).BackColor)
      
   Unload Me
End Sub

Private Sub butSelectColour_Click()
   
   On Error GoTo COLOUR_ERROR
   
   CommonDialog1.Color = picBackgroundSample(1).BackColor
   'CommonDialog1.Flags = cdlCCRGBInit Or cdlCCPreventFullOpen
   CommonDialog1.Flags = cdlCCRGBInit
   CommonDialog1.CancelError = True
   CommonDialog1.ShowColor
   picBackgroundSample(1).BackColor = CommonDialog1.Color
   
COLOUR_ERROR:

End Sub

Private Sub chkIsManyCartoonsDIsplayed_Click()
   If chkIsManyCartoonsDIsplayed.Value = 1 Then
      chkDisplayRefersh.Enabled = True
   Else
      chkDisplayRefersh.Enabled = False
   End If
End Sub

Private Sub cmbDesktopEffect_Click()
   Dim i1%

   For i1 = 0 To 2
      picBackgroundSample(i1).Visible = False
   Next
   If cmbDesktopEffect.ListIndex <> -1 Then
      picBackgroundSample(cmbDesktopEffect.ListIndex).Visible = True
   End If
   If cmbDesktopEffect.ListIndex = 1 Then
      butSelectColour.Visible = True
   Else
      butSelectColour.Visible = False
   End If
   
End Sub

Private Sub Form_Load()
   
   On Error Resume Next
   
   cmbCartoonInterval.AddItem "ten seconds"
   cmbCartoonInterval.AddItem "twenty seconds"
   cmbCartoonInterval.AddItem "thirty seconds"
   cmbCartoonInterval.AddItem "minute"
   cmbCartoonInterval.AddItem "two minutes"
   cmbCartoonInterval.AddItem "ten minutes"
   
   Select Case Trim(UCase(CStr(GetSetting(App.EXEName, "Configuration", "DisplayDuration", "20SEC"))))
      Case "10S", "10SEC", "0"
         cmbCartoonInterval.ListIndex = 0
      Case "20S", "20SEC", "1"
         cmbCartoonInterval.ListIndex = 1
      Case "30S", "30SEC", "2"
         cmbCartoonInterval.ListIndex = 2
      Case "1M", "1MIN", "3"
         cmbCartoonInterval.ListIndex = 3
      Case "2M", "2MIN", "4"
         cmbCartoonInterval.ListIndex = 4
      Case "10M", "10MIN", "5"
         cmbCartoonInterval.ListIndex = 5
      Case Else
         SaveSetting App.EXEName, "Configuration", "DisplayDuration", "20SEC"
         cmbCartoonInterval.ListIndex = 1
   End Select
         
   cmbDesktopEffect.AddItem "the desktop"
   cmbDesktopEffect.AddItem "cleared to this colour"
   cmbDesktopEffect.AddItem "a Dilbert-fest"
      
   Select Case Trim(UCase(CStr(GetSetting(App.EXEName, "Configuration", "BackdropEffect", "DESK"))))
      Case "DESK", "DESKTOP", "0"
         cmbDesktopEffect.ListIndex = 0
      Case "BLK", "BLANK", "1"
         cmbDesktopEffect.ListIndex = 1
      Case "DIL", "DILBERTS", "2"
         cmbDesktopEffect.ListIndex = 2
      Case Else
         SaveSetting App.EXEName, "Configuration", "BackdropEffect", "DESK"
         cmbDesktopEffect.ListIndex = 0
   End Select
      
   
   Select Case Trim(UCase(CStr(GetSetting(App.EXEName, "Configuration", "MultipleCartoons", "NO"))))
      Case "YES", "Y", "1"
         chkIsManyCartoonsDIsplayed.Value = vbChecked
      Case "NO", "N", "0"
         chkIsManyCartoonsDIsplayed.Value = vbUnchecked
      Case Else
         SaveSetting App.EXEName, "Configuration", "MultipleCartoons", "NO"
         chkIsManyCartoonsDIsplayed.Value = 0
   End Select
   
   Select Case Trim(UCase(CStr(GetSetting(App.EXEName, "Configuration", "MultipleRefresh", "NO"))))
      Case "YES", "Y", "1"
         chkDisplayRefersh.Value = vbChecked
      Case "NO", "N", "0"
         chkDisplayRefersh.Value = vbUnchecked
      Case Else
         SaveSetting App.EXEName, "Configuration", "MultipleRefresh", "NO"
         chkDisplayRefersh.Value = 0
   End Select
   
   If chkIsManyCartoonsDIsplayed.Value = vbChecked Then
      chkDisplayRefersh.Enabled = True
   Else
      chkDisplayRefersh.Enabled = False
   End If
   
   Select Case Trim(UCase(CStr(GetSetting(App.EXEName, "Configuration", "HideCursor", "NO"))))
      Case "YES", "Y", "1"
         chkHideCursor.Value = vbChecked
      Case "NO", "N", "0"
         chkHideCursor.Value = vbUnchecked
      Case Else
         SaveSetting App.EXEName, "Configuration", "HideCursor", "NO"
         chkHideCursor.Value = vbChecked
   End Select
   
   picBackgroundSample(1).BackColor = CLng("&H" & Trim(CStr(GetSetting(App.EXEName, "Configuration", "BackgroundColour", "00808080"))))
   
   If IsConfigParentWindowDefined Then
      tmrCheckParent.Enabled = True
   End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   labWeblink.ForeColor = &HFF0000
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then
      Me.Caption = "Hello?"
      tmrCaptionUpdate.Tag = 1
      tmrCaptionUpdate.Enabled = True
   Else
      tmrCaptionUpdate.Enabled = False
      tmrCaptionUpdate.Tag = 0
      Me.Caption = "Dilbert's On-Line Screen Saver Configuration"
   End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   labWeblink.ForeColor = &HFF0000
End Sub

Private Sub labWeblink_Click()
  
  Me.WindowState = vbMinimized
  
  ShellURL "http://www.dilbert.com"

End Sub

Private Sub labWeblink_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   labWeblink.ForeColor = &HFFFF00
End Sub

Private Sub tmrCaptionUpdate_Timer()
      
   Select Case tmrCaptionUpdate.Tag
      Case 0
         tmrCaptionUpdate.Enabled = False
         Me.Caption = "Dilbert's On-Line Screen Saver Configuration"
         Exit Sub
      Case 1
         Me.Caption = "It's me!"
      Case 2
         Me.Caption = "Dilbert!"
      Case 3
         Me.Caption = "Anyone there?"
      Case 4
         Me.Caption = "I'm too stressed..."
      Case 5
         Me.Caption = "...to ignore!"
      Case 6, 7, 8
         Me.Caption = ""
      Case 9
         Me.Caption = "Hello?"
         tmrCaptionUpdate.Tag = 0
   End Select
   
   tmrCaptionUpdate.Tag = tmrCaptionUpdate.Tag + 1
End Sub

Private Sub tmrCheckParent_Timer()
      
   On Error Resume Next
   
   If HasConfigParentWindowClosed Then
      tmrCheckParent.Enabled = False
      Unload Me
   End If
End Sub
