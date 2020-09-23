VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dilbert's On-Line Screen Saver Settings"
   ClientHeight    =   4095
   ClientLeft      =   3540
   ClientTop       =   2835
   ClientWidth     =   9330
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4095
   ScaleWidth      =   9330
   Visible         =   0   'False
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   1560
      Left            =   6225
      Picture         =   "frmSettings.frx":08CA
      ScaleHeight     =   1500
      ScaleWidth      =   2955
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picDesktop 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   1560
      Left            =   6195
      Picture         =   "frmSettings.frx":F36C
      ScaleHeight     =   1500
      ScaleWidth      =   2955
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picBackgroundSample 
      Enabled         =   0   'False
      Height          =   1560
      Index           =   5
      Left            =   6225
      Picture         =   "frmSettings.frx":1100A
      ScaleHeight     =   1500
      ScaleWidth      =   2955
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picBackgroundSample 
      Enabled         =   0   'False
      Height          =   1560
      Index           =   4
      Left            =   6225
      Picture         =   "frmSettings.frx":13E4F
      ScaleHeight     =   1500
      ScaleWidth      =   2955
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picBackgroundSample 
      Enabled         =   0   'False
      Height          =   1560
      Index           =   3
      Left            =   6225
      Picture         =   "frmSettings.frx":16D76
      ScaleHeight     =   1500
      ScaleWidth      =   2955
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picBackgroundSample 
      Enabled         =   0   'False
      Height          =   1560
      Index           =   2
      Left            =   6225
      Picture         =   "frmSettings.frx":1A0F0
      ScaleHeight     =   1500
      ScaleWidth      =   2955
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picBackgroundSample 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   1560
      Index           =   1
      Left            =   6195
      ScaleHeight     =   1500
      ScaleWidth      =   2955
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox chkRandomDisplay 
      Caption         =   "Display the cartoons in a &random order."
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   1530
      Width           =   3810
   End
   Begin VB.HScrollBar hscDesktopFade 
      Height          =   225
      LargeChange     =   32
      Left            =   7665
      Max             =   255
      SmallChange     =   8
      TabIndex        =   11
      Top             =   2190
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraCartoonInformation 
      Caption         =   "Cartoon Information"
      Height          =   2025
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   5355
      Begin VB.ComboBox cmbDownloadFinish 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1560
         Width           =   870
      End
      Begin VB.ComboBox cmbDownloadStart 
         Height          =   315
         Left            =   525
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1560
         Width           =   870
      End
      Begin VB.CheckBox chkRestrictDownload 
         Caption         =   "Restict &downloads to between..."
         Height          =   360
         Left            =   195
         TabIndex        =   13
         Top             =   1230
         Width           =   2655
      End
      Begin VB.CommandButton butDeleteCartoons 
         Caption         =   "Clear &archive..."
         Enabled         =   0   'False
         Height          =   390
         Left            =   3840
         TabIndex        =   16
         Top             =   1035
         Width           =   1365
      End
      Begin VB.CheckBox chkDisableDownload 
         Caption         =   "Disable dowloading from &website"
         Enabled         =   0   'False
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   1020
         Width           =   2670
      End
      Begin VB.Label labRestrictDownloadAnalysis 
         Height          =   240
         Left            =   2835
         TabIndex        =   41
         Top             =   1620
         Width           =   2010
      End
      Begin VB.Label labRestrictDownload 
         Caption         =   "and"
         Height          =   270
         Left            =   1440
         TabIndex        =   40
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label labLastDownloadPassive2 
         Caption         =   "ago."
         Height          =   225
         Left            =   3720
         TabIndex        =   30
         Top             =   585
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label labLastDownload 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2655
         TabIndex        =   29
         Top             =   570
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label labLastDownloadPassive1 
         Caption         =   "The last cartoon was downloaded"
         Height          =   225
         Left            =   180
         TabIndex        =   28
         Top             =   585
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label labNumDownloadsPassive2 
         Caption         =   "cartoons downloaded."
         Height          =   225
         Left            =   2160
         TabIndex        =   27
         Top             =   285
         Width           =   1995
      End
      Begin VB.Label labNumDownloads 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "no"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1605
         TabIndex        =   26
         Top             =   270
         Width           =   495
      End
      Begin VB.Label labNumDownloadsPassive1 
         Caption         =   "You currently have        "
         Height          =   225
         Left            =   195
         TabIndex        =   25
         Top             =   285
         Width           =   1335
      End
   End
   Begin VB.CommandButton butSelectColour 
      Caption         =   "&Colour..."
      Height          =   390
      Left            =   8100
      TabIndex        =   9
      Top             =   2175
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox picBackgroundSample 
      Enabled         =   0   'False
      Height          =   1560
      Index           =   0
      Left            =   6195
      Picture         =   "frmSettings.frx":1B16C
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
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
      ItemData        =   "frmSettings.frx":1D688
      Left            =   7260
      List            =   "frmSettings.frx":1D68A
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
      Left            =   8010
      TabIndex        =   18
      Top             =   3630
      Width           =   1245
   End
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6660
      TabIndex        =   17
      Top             =   3630
      Width           =   1245
   End
   Begin VB.ComboBox cmbCartoonInterval 
      Height          =   315
      ItemData        =   "frmSettings.frx":1D68C
      Left            =   3885
      List            =   "frmSettings.frx":1D68E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   165
      Width           =   1515
   End
   Begin VB.CheckBox chkDisplayRefersh 
      Caption         =   "but the &screen will be cleared every 10 minutes."
      Height          =   240
      Left            =   495
      TabIndex        =   6
      Top             =   870
      Width           =   3690
   End
   Begin VB.PictureBox picFade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   6195
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label labFullFade 
      Caption         =   "full"
      Height          =   210
      Left            =   8970
      TabIndex        =   34
      Top             =   2415
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label labNoFade 
      Caption         =   "none"
      Height          =   210
      Left            =   7710
      TabIndex        =   33
      Top             =   2415
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label labDesktopFade 
      Caption         =   "&Fade level"
      Height          =   225
      Left            =   6840
      TabIndex        =   10
      Top             =   2220
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label labCopyright 
      Caption         =   "Dilbert © 2003, United Feature Syndicate, Inc. "
      Height          =   270
      Left            =   5895
      TabIndex        =   31
      Top             =   3120
      Width           =   3555
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
      Left            =   6420
      TabIndex        =   23
      ToolTipText     =   "Dilbert Homepage"
      Top             =   2850
      Width           =   1305
   End
   Begin VB.Label labWeblinkPassive 
      Caption         =   "Visit the                                for more related fun!"
      Height          =   285
      Left            =   5775
      TabIndex        =   22
      Top             =   2850
      Width           =   3450
   End
   Begin VB.Label labDilbert 
      Caption         =   "Dilbert"
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   225
      Width           =   570
   End
   Begin VB.Label labDesktopEffect 
      Caption         =   "and the &backdrop will be"
      Height          =   255
      Left            =   5445
      TabIndex        =   3
      Top             =   210
      Width           =   1965
   End
   Begin VB.Image imaDilbert 
      Height          =   480
      Left            =   570
      Picture         =   "frmSettings.frx":1D690
      ToolTipText     =   "Hey! Get that cursor off my face, man!"
      Top             =   60
      Width           =   480
   End
   Begin VB.Label labCartoonInterval 
      Caption         =   "will show you a new cartoon strip &every"
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Top             =   225
      Width           =   3180
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const CC_RGBINIT = &H1

Private Declare Function ChooseColour Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Dim custom_colour_array() As Byte

Private backdrop_color&

Private Sub butDeleteCartoons_Click()
   frmClearArchive.Show vbModal, Me
   UpdateCartoonStats
End Sub

Private Sub chkDisableDownload_Click()

   If chkDisableDownload.Value = vbChecked Then
      chkRestrictDownload.Enabled = False
      cmbDownloadStart.Enabled = False
      cmbDownloadFinish.Enabled = False
      labRestrictDownload.Enabled = False
      labRestrictDownloadAnalysis.Enabled = False
   Else
      chkRestrictDownload.Enabled = True
      chkRestrictDownload_Click
   End If

End Sub

Private Sub chkRestrictDownload_Click()
   
   If chkRestrictDownload.Value = vbChecked Then
      cmbDownloadStart.Enabled = True
      cmbDownloadFinish.Enabled = True
      labRestrictDownload.Enabled = True
      labRestrictDownloadAnalysis.Enabled = True
   Else
      cmbDownloadStart.Enabled = False
      cmbDownloadFinish.Enabled = False
      labRestrictDownload.Enabled = False
      labRestrictDownloadAnalysis.Enabled = False
   End If

End Sub

Private Sub cmbDownloadFinish_Click()
   labRestrictDownloadAnalysis.Caption = UpdateDownloadAnalysis(cmbDownloadStart.ListIndex, cmbDownloadFinish.ListIndex)
End Sub

Private Sub cmbDownloadStart_Click()
   labRestrictDownloadAnalysis.Caption = UpdateDownloadAnalysis(cmbDownloadStart.ListIndex, cmbDownloadFinish.ListIndex)
End Sub

Private Sub Form_Load()
   Dim index%, custom_color_string$
   
   On Error Resume Next
   
   picFade.Left = Me.width
   picDesktop.Left = Me.width
   picTemp.Left = Me.width

   cmbCartoonInterval.AddItem "ten seconds"
   cmbCartoonInterval.AddItem "twenty seconds"
   cmbCartoonInterval.AddItem "thirty seconds"
   cmbCartoonInterval.AddItem "minute"
   cmbCartoonInterval.AddItem "two minutes"
   cmbCartoonInterval.AddItem "ten minutes"
   
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "DisplayDuration", "20SEC"))))
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
         SaveSetting App.Title, "Configuration", "DisplayDuration", "20SEC"
         cmbCartoonInterval.ListIndex = 1
   End Select
         
   cmbDesktopEffect.AddItem "the desktop"
   cmbDesktopEffect.AddItem "cleared to this colour"
   cmbDesktopEffect.AddItem "a Dilbert-fest"
      
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "BackdropEffect", "DESK"))))
      Case "DESK", "DESKTOP", "0"
         cmbDesktopEffect.ListIndex = 0
      Case "BLK", "BLANK", "1"
         cmbDesktopEffect.ListIndex = 1
      Case "DIL", "DILBERTS", "2"
         cmbDesktopEffect.ListIndex = 2
      Case Else
         SaveSetting App.Title, "Configuration", "BackdropEffect", "DESK"
         cmbDesktopEffect.ListIndex = 0
   End Select
      
   
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "MultipleCartoons", "NO"))))
      Case "YES", "Y", "1"
         chkIsManyCartoonsDIsplayed.Value = vbChecked
      Case "NO", "N", "0"
         chkIsManyCartoonsDIsplayed.Value = vbUnchecked
      Case Else
         SaveSetting App.Title, "Configuration", "MultipleCartoons", "NO"
         chkIsManyCartoonsDIsplayed.Value = vbUnchecked
   End Select
   
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "MultipleRefresh", "NO"))))
      Case "YES", "Y", "1"
         chkDisplayRefersh.Value = vbChecked
      Case "NO", "N", "0"
         chkDisplayRefersh.Value = vbUnchecked
      Case Else
         SaveSetting App.Title, "Configuration", "MultipleRefresh", "NO"
         chkDisplayRefersh.Value = vbUnchecked
   End Select
   
   If chkIsManyCartoonsDIsplayed.Value = vbChecked Then
      chkDisplayRefersh.Enabled = True
   Else
      chkDisplayRefersh.Enabled = False
   End If
   
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "HideCursor", "NO"))))
      Case "YES", "Y", "1"
         chkHideCursor.Value = vbChecked
      Case "NO", "N", "0"
         chkHideCursor.Value = vbUnchecked
      Case Else
         SaveSetting App.Title, "Configuration", "HideCursor", "NO"
         chkHideCursor.Value = vbChecked
   End Select
   
   backdrop_color = CLng("&H" & Trim(CStr(GetSetting(App.Title, "Configuration", "BackgroundColour", "00808080"))))
   picBackgroundSample(1).BackColor = backdrop_color
   
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "DisableDownloads", "NO"))))
      Case "YES", "Y", "1"
         chkDisableDownload.Value = vbChecked
      Case "NO", "N", "0"
         chkDisableDownload.Value = vbUnchecked
      Case Else
         SaveSetting App.Title, "Configuration", "DisableDownloads", "NO"
         chkDisableDownload.Value = vbUnchecked
   End Select

   UpdateCartoonStats
   
   custom_color_string = Left$(GetSetting(App.Title, "Configuration", "CustomColours", "") & String(128, "0"), 128)
   
   ReDim custom_colour_array(0 To 16 * 4 - 1) As Byte

   For index = LBound(custom_colour_array) To UBound(custom_colour_array)
      custom_colour_array(index) = CByte("&H" & Right$("00" & Mid$(custom_color_string, index * 2 + 1, 2), 2))
   Next
   
   hscDesktopFade.Value = CInt(GetSetting(App.Title, "Configuration", "DesktopFade", "128"))
         
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "RandomOrder", "NO"))))
      Case "YES", "Y", "1"
         chkRandomDisplay.Value = vbChecked
      Case "NO", "N", "0"
         chkRandomDisplay.Value = vbUnchecked
      Case Else
         SaveSetting App.Title, "Configuration", "RandomOrder", "NO"
         chkRandomDisplay.Value = vbUnchecked
   End Select
   
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "RestrictDownloads", "NO"))))
      Case "YES", "Y", "1"
         chkRestrictDownload.Value = vbChecked
      Case "NO", "N", "0"
         chkRestrictDownload.Value = vbUnchecked
      Case Else
         SaveSetting App.Title, "Configuration", "RestrictDownloads", "NO"
         chkRestrictDownload.Value = vbUnchecked
   End Select
   
   cmbDownloadStart.AddItem "12am"
   cmbDownloadStart.AddItem "1am"
   cmbDownloadStart.AddItem "2am"
   cmbDownloadStart.AddItem "3am"
   cmbDownloadStart.AddItem "4am"
   cmbDownloadStart.AddItem "5am"
   cmbDownloadStart.AddItem "6am"
   cmbDownloadStart.AddItem "7am"
   cmbDownloadStart.AddItem "8am"
   cmbDownloadStart.AddItem "9am"
   cmbDownloadStart.AddItem "10am"
   cmbDownloadStart.AddItem "11am"
   cmbDownloadStart.AddItem "12pm"
   cmbDownloadStart.AddItem "1pm"
   cmbDownloadStart.AddItem "2pm"
   cmbDownloadStart.AddItem "3pm"
   cmbDownloadStart.AddItem "4pm"
   cmbDownloadStart.AddItem "5pm"
   cmbDownloadStart.AddItem "6pm"
   cmbDownloadStart.AddItem "7pm"
   cmbDownloadStart.AddItem "8pm"
   cmbDownloadStart.AddItem "9pm"
   cmbDownloadStart.AddItem "10pm"
   cmbDownloadStart.AddItem "11pm"
   
   index = -1
   index = CInt(Trim(CStr(GetSetting(App.Title, "Configuration", "DownloadStart", "19"))))
   If index >= 0 And index < 24 Then
         cmbDownloadStart.ListIndex = index
   Else
         SaveSetting App.Title, "Configuration", "DownloadStart", "19"
         cmbDownloadStart.ListIndex = 19
   End If
   
   cmbDownloadFinish.AddItem "12am"
   cmbDownloadFinish.AddItem "1am"
   cmbDownloadFinish.AddItem "2am"
   cmbDownloadFinish.AddItem "3am"
   cmbDownloadFinish.AddItem "4am"
   cmbDownloadFinish.AddItem "5am"
   cmbDownloadFinish.AddItem "6am"
   cmbDownloadFinish.AddItem "7am"
   cmbDownloadFinish.AddItem "8am"
   cmbDownloadFinish.AddItem "9am"
   cmbDownloadFinish.AddItem "10am"
   cmbDownloadFinish.AddItem "11am"
   cmbDownloadFinish.AddItem "12pm"
   cmbDownloadFinish.AddItem "1pm"
   cmbDownloadFinish.AddItem "2pm"
   cmbDownloadFinish.AddItem "3pm"
   cmbDownloadFinish.AddItem "4pm"
   cmbDownloadFinish.AddItem "5pm"
   cmbDownloadFinish.AddItem "6pm"
   cmbDownloadFinish.AddItem "7pm"
   cmbDownloadFinish.AddItem "8pm"
   cmbDownloadFinish.AddItem "9pm"
   cmbDownloadFinish.AddItem "10pm"
   cmbDownloadFinish.AddItem "11pm"
   
   index = -1
   index = CInt(Trim(CStr(GetSetting(App.Title, "Configuration", "DownloadFinish", "7"))))
   If index >= 0 And index < 24 Then
         cmbDownloadFinish.ListIndex = index
   Else
         SaveSetting App.Title, "Configuration", "DownloadFinish", "7"
         cmbDownloadFinish.ListIndex = 7
   End If
      
   CentreSettingsForm Me
   
   If IsConfigParentWindowDefined Then
      tmrCheckParent.Enabled = True
   End If
   
End Sub

Private Sub Form_Paint()
   DoEvents
   hscDesktopFade_Change
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


Private Sub hscDesktopFade_Change()
   
   If UseSimpleFading Then
      Select Case hscDesktopFade.Value
         Case Is < 51
            picBackgroundSample(1).Visible = False
            picBackgroundSample(3).Visible = False
            picBackgroundSample(4).Visible = False
            picBackgroundSample(5).Visible = False
            picBackgroundSample(0).Visible = True
         Case 51 To 101
            picBackgroundSample(0).Visible = False
            picBackgroundSample(1).Visible = False
            picBackgroundSample(4).Visible = False
            picBackgroundSample(5).Visible = False
            picBackgroundSample(3).Visible = True
         Case 102 To 152
            picBackgroundSample(0).Visible = False
            picBackgroundSample(1).Visible = False
            picBackgroundSample(3).Visible = False
            picBackgroundSample(5).Visible = False
            picBackgroundSample(4).Visible = True
         Case 153 To 203
            picBackgroundSample(0).Visible = False
            picBackgroundSample(1).Visible = False
            picBackgroundSample(3).Visible = False
            picBackgroundSample(4).Visible = False
            picBackgroundSample(5).Visible = True
         Case Is > 203
            picBackgroundSample(0).Visible = False
            picBackgroundSample(3).Visible = False
            picBackgroundSample(4).Visible = False
            picBackgroundSample(5).Visible = False
            picBackgroundSample(1).BackColor = 0
            picBackgroundSample(1).Visible = True
      End Select
   Else
      AlphaBlendImageToImage picBackgroundSample(0), picDesktop, picFade, picTemp, hscDesktopFade.Value
   End If
   
End Sub


Private Sub labCopyright_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If labWeblink.ForeColor <> &HFF0000 Then
      labWeblink.ForeColor = &HFF0000
   End If

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

Private Sub butSelectColour_Click()
   Dim cc As CHOOSECOLOR

   On Error Resume Next

   cc.lStructSize = Len(cc)
   cc.hwndOwner = Me.hwnd
   cc.hInstance = App.hInstance
   cc.lpCustColors = StrConv(custom_colour_array, vbUnicode)
   cc.rgbResult = picBackgroundSample(1).BackColor
   cc.flags = CC_RGBINIT

   ' Show the 'Select Color' dialog.
   If ChooseColour(cc) <> 0 Then
      backdrop_color = cc.rgbResult
      picBackgroundSample(1).BackColor = backdrop_color
      custom_colour_array = StrConv(cc.lpCustColors, vbFromUnicode)
   End If
   
End Sub

Private Sub chkIsManyCartoonsDIsplayed_Click()
   
   If chkIsManyCartoonsDIsplayed.Value = vbChecked Then
      chkDisplayRefersh.Enabled = True
   Else
      chkDisplayRefersh.Enabled = False
   End If

End Sub

Private Sub cmbDesktopEffect_Click()
   Dim index%

   For index = 0 To 2
      picBackgroundSample(index).Visible = False
   Next
   
   Select Case cmbDesktopEffect.ListIndex
      Case 0
         picBackgroundSample(0).Visible = True
      Case 1
         picBackgroundSample(1).BackColor = backdrop_color
         picBackgroundSample(1).Visible = True
      Case 2
         picBackgroundSample(2).Visible = True
   End Select
   
   If cmbDesktopEffect.ListIndex <> -1 Then
      picBackgroundSample(cmbDesktopEffect.ListIndex).Visible = True
   End If
   
   If cmbDesktopEffect.ListIndex = 1 Then
      butSelectColour.Visible = True
   Else
      butSelectColour.Visible = False
   End If
   
   If cmbDesktopEffect.ListIndex = 0 Then
      labDesktopFade.Visible = True
      hscDesktopFade.Visible = True
      labFullFade.Visible = True
      labNoFade.Visible = True
      DoEvents
      hscDesktopFade_Change
   Else
      labDesktopFade.Visible = False
      hscDesktopFade.Visible = False
      labFullFade.Visible = False
      labNoFade.Visible = False
   End If
   
End Sub

Private Sub labWeblink_Click()
  
  Me.WindowState = vbMinimized
  
  ShellURL "http://www.dilbert.com"

End Sub

Private Sub labWeblink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   labWeblink.ForeColor = &HFFFF00
 
End Sub

Private Sub labWeblinkPassive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If labWeblink.ForeColor <> &HFF0000 Then
      labWeblink.ForeColor = &HFF0000
   End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If labWeblink.ForeColor <> &HFF0000 Then
      labWeblink.ForeColor = &HFF0000
   End If
   
End Sub


Private Sub butCancel_Click()
   
   Unload Me

End Sub

Private Sub butOk_Click()
   Dim index%, setting_string$, custom_colour_string$
   
   On Error Resume Next
   
   Select Case cmbCartoonInterval.ListIndex
      Case 0
         setting_string = "10SEC"
      Case 1
         setting_string = "20SEC"
      Case 2
         setting_string = "30SEC"
      Case 3
         setting_string = "1MIN"
      Case 4
         setting_string = "2MIN"
      Case 5
         setting_string = "10MIN"
      Case Else
         setting_string = "20SEC"
   End Select
      
   SaveSetting App.Title, "Configuration", "DisplayDuration", setting_string
   
   Select Case cmbDesktopEffect.ListIndex
      Case 0
         setting_string = "DESK"
      Case 1
         setting_string = "BLK"
      Case 2
         setting_string = "DIL"
      Case Else
         setting_string = "DESK"
   End Select
   
   SaveSetting App.Title, "Configuration", "BackdropEffect", setting_string
      
   If chkIsManyCartoonsDIsplayed.Value = vbChecked Then
      SaveSetting App.Title, "Configuration", "MultipleCartoons", "YES"
   Else
      SaveSetting App.Title, "Configuration", "MultipleCartoons", "NO"
   End If
   
   If chkDisplayRefersh.Value = vbChecked Then
      SaveSetting App.Title, "Configuration", "MultipleRefresh", "YES"
   Else
      SaveSetting App.Title, "Configuration", "MultipleRefresh", "NO"
   End If
   
   If chkHideCursor.Value = vbChecked Then
      SaveSetting App.Title, "Configuration", "HideCursor", "YES"
   Else
      SaveSetting App.Title, "Configuration", "HideCursor", "NO"
   End If
      
   If chkDisableDownload.Value = vbChecked And chkDisableDownload.Enabled = True Then
      SaveSetting App.Title, "Configuration", "DisableDownloads", "YES"
   Else
      SaveSetting App.Title, "Configuration", "DisableDownloads", "NO"
   End If
      
   SaveSetting App.Title, "Configuration", "BackgroundColour", Hex(backdrop_color)
      
   For index = LBound(custom_colour_array) To UBound(custom_colour_array)
      custom_colour_string = custom_colour_string & Right$("00" & Hex$(custom_colour_array(index)), 2)
   Next
   
   SaveSetting App.Title, "Configuration", "CustomColours", custom_colour_string
   
   SaveSetting App.Title, "Configuration", "DesktopFade", hscDesktopFade.Value
   
   If chkRandomDisplay.Value = vbChecked Then
      SaveSetting App.Title, "Configuration", "RandomOrder", "YES"
   Else
      SaveSetting App.Title, "Configuration", "RandomOrder", "NO"
   End If
   
   If chkRestrictDownload.Value = vbChecked Then
      SaveSetting App.Title, "Configuration", "RestrictDownloads", "YES"
   Else
      SaveSetting App.Title, "Configuration", "RestrictDownloads", "NO"
   End If
   
   SaveSetting App.Title, "Configuration", "DownloadStart", CStr(cmbDownloadStart.ListIndex)
   
   SaveSetting App.Title, "Configuration", "DownloadFinish", CStr(cmbDownloadFinish.ListIndex)
   
   Unload Me

End Sub

Private Sub UpdateCartoonStats()
   Dim last_download_date As Date, last_downlaod_days_ago&
   
   labNumDownloads.Caption = Trim(CStr(GetSetting(App.Title, "Cartoons", "NumDownloads", "no")))
   If labNumDownloads.Caption <> "no" Then
      labLastDownload.Visible = True
      labLastDownloadPassive1.Visible = True
      labLastDownloadPassive2.Visible = True
      chkDisableDownload.Enabled = True
      butDeleteCartoons.Enabled = True
      last_download_date = CDate(Trim(CStr(GetSetting(App.Title, "Cartoons", "LastDownload", CStr(Now)))))
      last_downlaod_days_ago = DateDiff("d", last_download_date, Now)
      
      If last_downlaod_days_ago > 90 Then
         labLastDownload.Caption = CStr(last_downlaod_days_ago \ 30) & " months"
      ElseIf last_downlaod_days_ago = 60 Then
         labLastDownload.Caption = "2 months"
      ElseIf last_downlaod_days_ago = 30 Then
         labLastDownload.Caption = "1 month"
      ElseIf last_downlaod_days_ago > 21 Then
         labLastDownload.Caption = CStr(last_downlaod_days_ago \ 7) & " weeeks"
      ElseIf last_downlaod_days_ago = 14 Then
         labLastDownload.Caption = "2 weeeks"
      ElseIf last_downlaod_days_ago = 7 Then
         labLastDownload.Caption = "1 weeek"
      ElseIf last_downlaod_days_ago = 1 Then
         labLastDownload.Caption = "yesterday"
         labLastDownloadPassive2.Visible = False
      ElseIf last_downlaod_days_ago = 0 Then
         labLastDownload.Caption = "today"
         labLastDownloadPassive2.Visible = False
      Else
         labLastDownload.Caption = CStr(last_downlaod_days_ago) & " days"
      End If
   End If
End Sub

Static Function UpdateDownloadAnalysis(ByVal start_index%, ByVal finish_index%) As String
   Dim diff%
   
   On Error Resume Next
   
   UpdateDownloadAnalysis = ""
   
   If start_index = -1 Or finish_index = -1 Then
      Exit Function
   End If
   
   If start_index < finish_index Then
      diff = finish_index - start_index
      UpdateDownloadAnalysis = "(" & CStr(diff) & " hour" & IIf(diff = 1, "", "s") & ", "
      If diff <= 12 Then
         If finish_index <= 8 Then
            UpdateDownloadAnalysis = UpdateDownloadAnalysis & "early morning)"
         ElseIf start_index >= 18 Then
            UpdateDownloadAnalysis = UpdateDownloadAnalysis & "evening)"
         Else
            UpdateDownloadAnalysis = UpdateDownloadAnalysis & "daytime)"
         End If
      Else
         UpdateDownloadAnalysis = UpdateDownloadAnalysis & "over day)"
      End If
   ElseIf finish_index < start_index Then
      diff = 24 - (start_index - finish_index)
      UpdateDownloadAnalysis = "(" & CStr(diff) & " hour" & IIf(diff = 1, "", "s")
      If diff <= 12 Then
         UpdateDownloadAnalysis = UpdateDownloadAnalysis & ", overnight)"
      Else
         UpdateDownloadAnalysis = UpdateDownloadAnalysis & ")"
      End If
   Else
      UpdateDownloadAnalysis = "(single attempt)"
   End If

End Function
