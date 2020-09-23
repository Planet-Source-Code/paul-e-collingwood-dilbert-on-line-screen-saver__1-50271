VERSION 5.00
Begin VB.Form frmClearArchive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clear Cartoon Archive"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmClearArchive.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2025
      TabIndex        =   7
      Top             =   1110
      Width           =   1245
   End
   Begin VB.CommandButton butOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   3375
      TabIndex        =   6
      Top             =   1110
      Width           =   1245
   End
   Begin VB.OptionButton optClear 
      Caption         =   "Clear all cartoons"
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   2
      Top             =   780
      Width           =   2985
   End
   Begin VB.ComboBox cmbLeaveCount 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   375
      Width           =   930
   End
   Begin VB.OptionButton optClear 
      Caption         =   "Leave latest                        cartoons."
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   420
      Width           =   2985
   End
   Begin VB.Label labNumDownloadsPassive1 
      Caption         =   "You currently have        "
      Height          =   225
      Left            =   135
      TabIndex        =   5
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label labNumDownloads 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "no"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1545
      TabIndex        =   4
      Top             =   45
      Width           =   495
   End
   Begin VB.Label labNumDownloadsPassive2 
      Caption         =   "cartoons downloaded."
      Height          =   225
      Left            =   2100
      TabIndex        =   3
      Top             =   60
      Width           =   1995
   End
End
Attribute VB_Name = "frmClearArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub butCancel_Click()
   Unload Me
End Sub

Private Sub butOk_Click()
   Dim temporary_path$, file_name$, file_prefix$, file_count&, files_required&

   On Error Resume Next

   butOk.Enabled = False
   butCancel.Enabled = False

   temporary_path = GetCartoonFileArchivePath()
   file_count = 0
   
   file_name = Dir(temporary_path & "\*.*", vbNormal)
   Do While file_name <> ""
      file_prefix = UCase(Right$(file_name, 4))
      If file_prefix = ".GIF" Or file_prefix = ".JPG" Then
         file_count = file_count + 1
      End If

      file_name = Dir
   Loop

   If optClear(0).Value = True Then
      files_required = CLng(cmbLeaveCount.List(cmbLeaveCount.ListIndex))
   Else
      files_required = 0
   End If

   file_name = Dir(temporary_path & "\*.*", vbNormal)
   Do While file_name <> ""
      file_prefix = UCase(Right$(file_name, 4))
      If file_prefix = ".GIF" Or file_prefix = ".JPG" Then
         If files_required >= file_count Then
            Exit Do
         End If
         Kill temporary_path & "\" & file_name
         
         file_count = file_count - 1
      End If

      file_name = Dir
   Loop
   
   If file_count = 0 Then
      SaveSetting App.Title, "Cartoons", "NumDownloads", "no"
   Else
      SaveSetting App.Title, "Cartoons", "NumDownloads", CStr(file_count)
   End If
   
   Unload Me
End Sub

Private Sub Form_Load()
   
   On Error Resume Next
   
   labNumDownloads.Caption = Trim(CStr(GetSetting(App.Title, "Cartoons", "NumDownloads", "no")))
   cmbLeaveCount.AddItem "5"
   cmbLeaveCount.AddItem "10"
   cmbLeaveCount.AddItem "20"
   cmbLeaveCount.AddItem "50"
   cmbLeaveCount.AddItem "100"
   cmbLeaveCount.AddItem "200"
   cmbLeaveCount.AddItem "500"
   cmbLeaveCount.AddItem "1000"
   cmbLeaveCount.AddItem "2000"
   cmbLeaveCount.ListIndex = 4
   optClear(0).Value = True
End Sub
