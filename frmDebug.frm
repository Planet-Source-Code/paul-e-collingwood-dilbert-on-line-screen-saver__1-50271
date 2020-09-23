VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chose the running mode"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHandle 
      Height          =   285
      Left            =   5070
      TabIndex        =   3
      Text            =   "0"
      Top             =   1485
      Width           =   1065
   End
   Begin VB.CommandButton butPreview 
      Caption         =   "Window Preview"
      Height          =   600
      Left            =   105
      TabIndex        =   2
      Top             =   1260
      Width           =   1065
   End
   Begin VB.CommandButton butConfigure 
      Caption         =   "Configure"
      Height          =   375
      Left            =   105
      TabIndex        =   1
      Top             =   765
      Width           =   1065
   End
   Begin VB.CommandButton butNormal 
      Caption         =   "Run / Preview"
      Default         =   -1  'True
      Height          =   600
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   1065
   End
   Begin VB.Label Label5 
      Caption         =   "You have to place the handle to a parent dialog in the TextBox - use a Windows Spy application to determine a valid setting here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1230
      TabIndex        =   8
      Top             =   1395
      Width           =   3780
   End
   Begin VB.Label Label4 
      Caption         =   "Click here to test the preview splash screen."
      Height          =   225
      Left            =   1260
      TabIndex        =   7
      Top             =   1170
      Width           =   3525
   End
   Begin VB.Label Label3 
      Caption         =   "Handle"
      Height          =   180
      Left            =   5085
      TabIndex        =   6
      Top             =   1275
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "Click here to test the settings dialog."
      Height          =   210
      Left            =   1275
      TabIndex        =   5
      Top             =   855
      Width           =   2850
   End
   Begin VB.Label Label1 
      Caption         =   "Click here to test the screen saver."
      Height          =   435
      Left            =   1260
      TabIndex        =   4
      Top             =   255
      Width           =   2835
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub butAbort_Click()
   End
End Sub

Private Sub butConfigure_Click()
   DebugSetConfigurationMode
   Me.Hide
   Unload Me
End Sub

Private Sub butNormal_Click()
   DebugSetScreensaverMode
   Me.Hide
   Unload Me
End Sub

Private Sub butPreview_Click()
   DebugSetPreviewMode "/p:" & txtHandle.Text
   Me.Hide
   Unload Me
End Sub

Private Sub txtHandle_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 32 Then
      If KeyAscii < Asc("0") Then
         KeyAscii = 0
      ElseIf KeyAscii > Asc("9") Then
         KeyAscii = 0
      End If
   End If
End Sub
