VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   2010
   ClientTop       =   2430
   ClientWidth     =   6375
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picHatch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   5355
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   2520
      Width           =   450
   End
   Begin VB.PictureBox picHatch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   5355
      Picture         =   "frmMain.frx":1D0C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   6
      Top             =   2017
      Width           =   450
   End
   Begin VB.PictureBox picHatch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   5355
      Picture         =   "frmMain.frx":314E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   1515
      Width           =   450
   End
   Begin VB.PictureBox picScreenFade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   4530
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   1095
      Width           =   450
   End
   Begin VB.Timer tmrCollate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5715
      Top             =   570
   End
   Begin VB.PictureBox picBackdrop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   4605
      Picture         =   "frmMain.frx":4590
      ScaleHeight     =   1155
      ScaleWidth      =   1080
      TabIndex        =   3
      Top             =   3060
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox picSplash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   0
      Picture         =   "frmMain.frx":59D2
      ScaleHeight     =   4245
      ScaleWidth      =   4470
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4470
   End
   Begin VB.Timer tmrDisplay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5685
      Top             =   120
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   4545
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   555
      Width           =   450
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   4545
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   30
      Width           =   450
   End
   Begin VB.Image imaSource 
      Height          =   450
      Left            =   4530
      Top             =   1635
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The various states of the cartoon display timer.
Private Const STATE_INIT = 0
Private Const STATE_STARTSCREEN = 1
Private Const STATE_DELAY = 2
Private Const STATE_SHOWCARTOON = 3
Private Const STATE_PREVIEW = 50

' Constants used to shift Preview image around (using polar rotation functions).
Private Const PREVIEW_ANGLE_DIFF1 = 0.2
Private Const PREVIEW_ANGLE_DIFF2 = 0.25
Private Const PREVIEW_ANGLE_DIFF3 = 0.06

' Main counter used in tmrDisplay to determine when durations have expired
' (has one-tenth second resolution).
Private timer_count&
Private cartoon_display_duration&
Private backdrop_effect_type%
Private cartoon_x_position%, cartoon_y_position%
Private input_exit_enabled As Boolean
Private preview_angle1#, preview_angle2#, preview_angle3#
Private multiple_cartoons_displayed As Boolean
Private multiple_refresh As Boolean
Private random_order As Boolean
Private refresh_count%

' Flag used to interrupt timer processes after UnLoad has been requested.
Private is_unloading As Boolean

' Variables used to hold and access the cartoon file list.
Dim cartoon_list() As String
Private cartoon_count&
Private next_cartoon_index&

Dim cartoon_archive_path$

Private Sub Form_Load()
   Dim index%
   
   On Error Resume Next
   
   cartoon_archive_path = GetCartoonFileArchivePath()
   
   MkDir cartoon_archive_path
   
   ' Reset cartoon list.
   cartoon_count = 0
   
   is_unloading = False
   input_exit_enabled = False
         
   ' Set the random seed differently each time the screen saver runs - this ensures
   ' different sequences of cartoon positions on the screen.
   Randomize Second(Now)
   
   ' Resize and position this form to cover the entire desktop display.
   Me.width = Screen.width
   Me.height = Screen.height
   Me.Left = Me.width + 100
   Me.Top = Me.height + 100
        
   ' Move the PictureBox controls used to buffer cartoon images to an off-screen position.
   picSource.Left = Me.width
   picSave.Left = Me.width
   picScreenFade.Left = Me.width
   picScreenFade.width = Screen.width \ Screen.TwipsPerPixelX
   picScreenFade.height = Screen.height \ Screen.TwipsPerPixelY
   For index = 0 To 2
      picHatch(index).Left = Me.width
   Next
   
   If RunningAsScreensaver Then
      ' Initiate the collation of details of all cartoon files on the hard-disk, and then download this month's
      ' Dilbert cartoons from the website (if rerquired).
      tmrCollate.Enabled = True
      
      ' Set the display state machine to run as a screen saver.
      tmrDisplay.Tag = STATE_INIT
               
      ' Grab snapshot of desktop screen area and copy to the screensaver form.
      CopyDesktopToForm Me
   Else
      frmMain.Caption = "Dilbert Screen Saver Preview"
      
      ' Set the display state machine to run as a preview window.
      tmrDisplay.Tag = STATE_PREVIEW
   End If
      
      ' Initiate display state machine timer.
   tmrDisplay.Enabled = True

End Sub

Private Sub Form_Click()

    If RunningAsScreensaver Then
      ' If a mouse click occurs then stop the screensaver.
      DoUnload
   End If

End Sub

Private Sub Form_DblClick()
    
    If RunningAsScreensaver Then
      ' If a mouse double-click occurs then stop the screensaver.
      DoUnload
   End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If RunningAsScreensaver Then
      ' If a key press occurs then stop the screensaver.
      DoUnload
   End If

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
   
   If (RunningAsScreensaver) And (input_exit_enabled = True) Then
      ' If a key press occurs then stop the screensaver.
      DoUnload
   End If

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If RunningAsScreensaver And (input_exit_enabled = True) Then
      ' If a mouse click occurs then stop the screensaver.
      DoUnload
   End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Static previous_x%, previous_y%

   If RunningAsScreensaver Then
       If ((previous_x = 0) And (previous_y = 0)) Or _
          ((Abs(previous_x - X) < 5) And (Abs(previous_y - Y) < 5)) Then
          ' A small movement occured, so dont bother to stop the screensaver.
          previous_x = X
          previous_y = Y
          Exit Sub
       End If
       
       ' Unload on large mouse movements.
       If input_exit_enabled = True Then
         DoUnload
      End If
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   On Error Resume Next
   
   If RunningAsScreensaver Then
      ' Redisplay the cursor.
      UnHideCursor
   End If
   
End Sub

Private Sub tmrCollate_Timer()
   Dim html_buffer$, html_buffer2$, day_search_string$, pos&, pos2&, pos3&, pos4&
   Dim date_string$, source_url$, file_name$, source_search_string$, file_prefix$
   Dim file_number%, first_found As Boolean, last_completed_download_date As Date
   Dim start_hour%, finish_hour%, the_hour%
   Dim data_buffer() As Byte
   
   On Error Resume Next
   
   ' This is a songle-shot process, so permanenntly disable this Timer.
   tmrCollate.Enabled = False
      
   ' Get array list of all relevant cartoon files from the relevant directory.
   file_name = Dir(cartoon_archive_path & "\*.*", vbNormal)
   Do While file_name <> ""
      file_prefix = UCase(Right$(file_name, 4))
      If file_prefix = ".GIF" Or file_prefix = ".JPG" Then
         Err.Clear
         ReDim Preserve cartoon_list(cartoon_count + 1) As String
         
         If Err.Number = 0 Then
            cartoon_list(cartoon_count) = file_name
            cartoon_count = cartoon_count + 1
         End If
      End If
      
      file_name = Dir
   Loop
   
   If cartoon_count = 0 Then
      SaveSetting App.Title, "Cartoons", "NumDownloads", "no"
   Else
      SaveSetting App.Title, "Cartoons", "NumDownloads", CStr(cartoon_count)
   End If
   
   'If there are more than one week's worth of cartoons, then start showing the past weeks first.
   ' NB - this is irrelevant if the screensaver is configured to display the cartoons randomly.
   If cartoon_count > 7 Then
      next_cartoon_index = cartoon_count - 7
   Else
      next_cartoon_index = 0
   End If
   
   ' Check to see if the screensaver has been halted - abort this code if so.
   If is_unloading = True Then
      Unload Me
      Exit Sub
   End If
   
   ' If configured to disable website downloads of more cartoons, then there is nothing left to do, so
   ' abort this process.
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "DisableDownloads", "NO"))))
      Case "YES", "Y", "1"
         Exit Sub
   End Select

   ' If a full download cycle is completed (i.e. checked that the month's worth of cartoons
   ' archived on the website,) already today, then there is no need to do it again, so
   ' abort this process.
   date_string = Trim(CStr(GetSetting(App.Title, "Cartoons", "LastCompletedDownload", "?")))
   If date_string <> "?" Then
      last_completed_download_date = CDate(date_string)
      
      If DateDiff("d", last_completed_download_date, Now) = 0 Then
         Exit Sub
      End If
   End If

   ' If configured to restrict website downloads of more cartoons between cratin hours of the day, then
   ' check to see that we are in the correct hourly time-band. If not, then idle this routine until we
   ' reach the required time. This, of course may quit if the screensaver is stopped!
   Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "RestrictDownloads", "NO"))))
      Case "YES", "Y", "1"
         ' Get starting hour (in 24-hour notation) of restricted download period (default is 7 pm)
         start_hour = -1
         start_hour = CInt(Trim(CStr(GetSetting(App.Title, "Configuration", "DownloadStart", "19"))))
         If start_hour < 0 Or start_hour > 23 Then
               start_hour = 19
         End If
         
         ' Get finishing hour (in 24-hour notation) of restricted download period (default is 7 am)
         finish_hour = -1
         finish_hour = CInt(Trim(CStr(GetSetting(App.Title, "Configuration", "DownloadFinish", "7"))))
         If finish_hour < 0 Or finish_hour > 23 Then
               finish_hour = 7
         End If
         
         ' If the hours given are the same, then advance finish by one hour - gives small duration!
         If finish_hour = start_hour Then
            If finish_hour = 23 Then
               finish_hour = 0
            Else
               finish_hour = start_hour + 1
            End If
         End If
         
         ' Now check to see if we are within hours given (and idle until we are!)
         If start_hour < finish_hour Then
            the_hour = Hour(Now)
            Do While the_hour < start_hour Or the_hour > finish_hour
               DoEvents
               
               ' Check to see if the screensaver has been halted - abort this code if so.
               If is_unloading = True Then
                  Unload Me
                  Exit Sub
               End If
               the_hour = Hour(Now)
            Loop
            
         ElseIf start_hour > finish_hour Then
            
            the_hour = Hour(Now)
            Do While the_hour < start_hour And the_hour > finish_hour
               DoEvents
               
               ' Check to see if the screensaver has been halted - abort this code if so.
               If is_unloading = True Then
                  Unload Me
                  Exit Sub
               End If
               the_hour = Hour(Now)
            Loop
         End If
   End Select


   first_found = True

   ' Get HTML content of Dilbert website's cartoon archive page.
   DownloadURLToString "www.dilbert.com", "/comics/dilbert/archive/index.html", html_buffer
      
   ' Check to see if the screensaver has been halted - abort this code if so.
   If is_unloading = True Then
      Unload Me
      Exit Sub
      End If
   
   ' Look for hyperlinks to each page of previous month - each one will reveal the URL of the actual
   ' cartoon image source.
   day_search_string = "<OPTION VALUE=" & Chr(34) & "/comics/dilbert/archive/dilbert-"
   pos = InStr(html_buffer, day_search_string)
   Do While pos <> 0
      DoEvents
      
      ' Check to see if the screensaver has been halted - abort this code if so.
      If is_unloading = True Then
         Unload Me
         Exit Sub
      End If
      
      ' Each hyperlink is to a page whose title is an encoding of the date (at present).
      ' This will extract the prefix component (??????.HTML).
      pos2 = InStr(pos, html_buffer, ".html" & Chr(34) & ">")
      If pos2 <> 0 Then
         date_string = Mid(html_buffer, pos + Len(day_search_string), pos2 - pos - Len(day_search_string))
            
         ' Get HTML content of each Dilbert website's cartoon page.
         html_buffer2 = ""
         DownloadURLToString "www.dilbert.com", "/comics/dilbert/archive/dilbert-" & date_string & ".html", html_buffer2
                  
         ' Check to see if the screensaver has been halted - abort this code if so.
         If is_unloading = True Then
            Unload Me
            Exit Sub
         End If
      
         ' Now extract the specific cartoons source URL - this will end in a .GIF or .JPG extension.
         source_search_string = "<IMG SRC=" & Chr(34) & "/comics/dilbert/archive/images/dilbert"
         pos2 = InStr(html_buffer2, source_search_string)
         If pos2 <> 0 Then
            pos3 = InStr(pos2, html_buffer2, ".gif" & Chr(34))
            If pos3 <> 0 Then
               pos4 = InStr(pos2, html_buffer2, ".jpg" & Chr(34))
               If pos4 <> 0 And pos4 < pos3 Then
                  source_url = Mid(html_buffer2, pos2 + Len(source_search_string), pos4 - pos2 - Len(source_search_string) + 4)
               Else
                  source_url = Mid(html_buffer2, pos2 + Len(source_search_string), pos3 - pos2 - Len(source_search_string) + 4)
               End If
               
               ' Now check to see if identical file exists in the archive on the hard-disk.
               file_name = cartoon_archive_path & "\" & source_url
               If Dir(file_name, vbNormal) = "" Then
                  Debug.Print source_url
                                       
                  ' Check to see if the screensaver has been halted - abort this code if so.
                  If is_unloading = True Then
                     Unload Me
                     Exit Sub
                  End If
               
                  ' If the caartoom image file is not already in our archive then download it.
                  If DownloadURLToFile("www.dilbert.com", "/comics/dilbert/archive/images/dilbert" & source_url, file_name) = True Then
                  
                     'If successfully downloaded, it is then now appended it to the array list.
                     Err.Clear
                     ReDim Preserve cartoon_list(cartoon_count + 1) As String
                     If Err.Number = 0 Then
                        cartoon_list(cartoon_count) = source_url
                        
                        If first_found = True Then
                           next_cartoon_index = cartoon_count
                           first_found = False
                        End If
                        
                        cartoon_count = cartoon_count + 1
                        
                        ' Update the relevant registry records.
                        SaveSetting App.Title, "Cartoons", "LastDownload", CStr(Now)
                        SaveSetting App.Title, "Cartoons", "NumDownloads", CStr(cartoon_count)

                     End If
                  End If
               End If
            End If
         End If
      End If
      
      pos = InStr(pos + 1, html_buffer, day_search_string)
   Loop

   ' Have sucessfully checked for all cartoons in the archive - by updating this registry record
   ' any subsequent checks of the website are prevented whenever the screensaver is run today.
   SaveSetting App.Title, "Cartoons", "LastCompletedDownload", CStr(Now)

   ' Check to see if the screensaver has been halted - abort this code if so.
   If is_unloading = True Then
      Unload Me
      Exit Sub
   End If
   
End Sub

Private Sub tmrDisplay_Timer()
   Dim i1%, i2%
   Dim sx%, sy%, cx%, cy%
   Dim dist As Double
   Dim fade_factor&
   
   On Error Resume Next
   
   ' This timer is continiuously running, and so this counter acts as an iteration count.
   ' Because of the Timer interval, it is updated no more than ten times per second.
   refresh_count = refresh_count + 1
   
   ' The Timer's Tag field is used to hold the state of the display cycle....
   Select Case tmrDisplay.Tag
      Case STATE_INIT
         ' Initialisation state - collate all relevant settings from registry records.
         Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "DisplayDuration", "20SEC"))))
            Case "10S", "10SEC", "0"
               cartoon_display_duration = 100
            Case "20S", "20SEC", "1"
               cartoon_display_duration = 200
            Case "30S", "30SEC", "2"
               cartoon_display_duration = 300
            Case "1M", "1MIN", "3"
               cartoon_display_duration = 600
            Case "2M", "2MIN", "4"
               cartoon_display_duration = 1200
            Case "10M", "10MIN", "5"
               cartoon_display_duration = 6000
            Case Else
               SaveSetting App.Title, "Configuration", "DisplayDuration", "5SEC"
               cartoon_display_duration = 200
         End Select
            
         Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "BackdropEffect", "DESK"))))
            Case "DESK", "DESKTOP", "0"
               backdrop_effect_type = 0
            Case "BLK", "BLANK", "1"
               backdrop_effect_type = 1
            Case "DIL", "DILBERTS", "2"
               backdrop_effect_type = 2
            Case Else
               SaveSetting App.Title, "Configuration", "BackdropEffect", "DESK"
               backdrop_effect_type = 0
         End Select
      
         Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "MultipleCartoons", "NO"))))
            Case "YES", "Y", "1"
               multiple_cartoons_displayed = True
            Case "NO", "N", "0"
               multiple_cartoons_displayed = False
            Case Else
               SaveSetting App.Title, "Configuration", "MultipleCartoons", "NO"
               multiple_cartoons_displayed = False
         End Select
      
         Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "MultipleRefresh", "NO"))))
            Case "YES", "Y", "1"
               If multiple_cartoons_displayed = True Then
                  multiple_refresh = True
               Else
                  multiple_refresh = False
               End If
            Case "NO", "N", "0"
               multiple_refresh = False
            Case Else
               SaveSetting App.Title, "Configuration", "MultipleRefresh", "NO"
               multiple_refresh = False
         End Select
                        
         Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "RandomOrder", "NO"))))
            Case "YES", "Y", "1"
               random_order = True
            Case "NO", "N", "0"
               random_order = False
            Case Else
               SaveSetting App.Title, "Configuration", "RandomOrder", "NO"
               random_order = False
         End Select
                        
         fade_factor = CInt(GetSetting(App.Title, "Configuration", "DesktopFade", "128"))
   
         ' Now update the entire display according to our required backdrop effect.
         ' NB - the form has been resizes to full screen, so this is not actually altering the desktop display at all.
         Select Case backdrop_effect_type
            Case 0
               ' Copy the entire desktop image to the screensaver form - this gives the illusion that catoons are
               ' displayed 'floating over other application windows etc.
               If UseSimpleFading Then
                  Select Case fade_factor
                     Case Is < 51
                        'Screen fade is off, so nothing to do
                     Case 51 To 101
                        TileImageToImage picScreenFade, picHatch(0)
                        AndImageToForm Me, picScreenFade
                     Case 102 To 152
                        TileImageToImage picScreenFade, picHatch(1)
                        AndImageToForm Me, picScreenFade
                     Case 153 To 203
                        TileImageToImage picScreenFade, picHatch(2)
                        AndImageToForm Me, picScreenFade
                     Case Is > 203
                        'Screen fade is total, so just black-out the desktop
                        frmMain.BackColor = 0
                        frmMain.Cls
                  End Select
               Else
                  AlphaBlendImageToForm Me, picScreenFade, fade_factor
               End If
            Case 1
               ' Clear our screensaver display to a specified colour plane.
               frmMain.BackColor = CLng("&H" & Trim(CStr(GetSetting(App.Title, "Configuration", "BackgroundColour", "00808080"))))
               frmMain.Cls
            Case 2
               ' Clear our screensaver display to black and paint lots of Dilbert motifs.
               frmMain.BackColor = 0
               frmMain.Cls
               sx = 0
               sy = -600
               For i1 = 0 To 12
                  cx = sx
                  cy = sy
                  For i2 = 0 To 25
                     CopyImageToForm Me, cx, cy, 64, 64, picBackdrop
                     cx = cx + 140
                     cy = cy + 60
                  Next
                  sx = sx - 300
                  sy = sy + 50
               Next
         End Select
         
         refresh_count = 0
         
         ' Reposition the screensaver form to cover the entire desktop area.
         Me.Left = 0
         Me.Top = 0
         ' Forcing it to be the foremost window also hides the TaskBar and System Tray.
         SetWindowTopmost Me
            
         ' Set the cartoon coordinates to a position off the display.
         ' Thus is done to effectively disable the first attempt to restore a saved area of the desktop display.
         cartoon_x_position = Screen.width / Screen.TwipsPerPixelX
         cartoon_y_position = Screen.height / Screen.TwipsPerPixelY
            
         timer_count = 0
         tmrDisplay.Tag = STATE_STARTSCREEN
         
      Case STATE_STARTSCREEN
         ' Post-diplay initialisation state - waits for screen display to be updated with screensaver form
         '                                    in correct position before allowing the application to be
         '                                    aborted by mouse activity.
         timer_count = timer_count + 1
         If timer_count >= 2 Then
            timer_count = cartoon_display_duration
            tmrDisplay.Tag = STATE_DELAY
            input_exit_enabled = True
         End If
         
      Case STATE_SHOWCARTOON
         frmMain.AutoRedraw = False
 
         If multiple_cartoons_displayed = False Then
            CopyImageFromForm Me, cartoon_x_position, cartoon_y_position, imaSource.width, imaSource.height, picSave
         End If
 
         CopyImageToForm Me, cartoon_x_position, cartoon_y_position, imaSource.width, imaSource.height, picSource
         frmMain.AutoRedraw = True
         
         timer_count = 0
         
         tmrDisplay.Tag = STATE_DELAY
            
      Case STATE_DELAY
         ' General display idle state - waits for configured duration and then prepares for the display of the next
         '                              cartoon. Optioanlly clears the display if required.
         timer_count = timer_count + 1
         If timer_count >= cartoon_display_duration Then
            
            ' If ten minutes have passed, then clear all cartoons from the display - if required
            If refresh_count > 6000 Then
               Select Case backdrop_effect_type
                  Case 0
                     ' Moving the screensaver form off and onto the display instantaneously clears of all cartoons from
                     ' the desktop image.
                     Me.Left = Screen.width
                     Me.Left = 0
                  Case 1
                     ' Clear our screensaver display to a specified colour plane.
                     frmMain.Cls
                  Case 2
                     ' Clear our screensaver display to black and paint lots of Dilbert motifs.
                     frmMain.BackColor = 0
                     frmMain.Cls
                     sx = 0
                     sy = -600
                     For i1 = 0 To 12
                        cx = sx
                        cy = sy
                        For i2 = 0 To 25
                           CopyImageToForm Me, cx, cy, 64, 64, picBackdrop
                           cx = cx + 140
                           cy = cy + 60
                        Next
                        sx = sx - 300
                        sy = sy + 50
                     Next
               End Select
                              
               refresh_count = 0
            Else
               ' If configured to allow only single cartoon on the display, then the copy of the background saved
               ' before the last cartoon was displayed is restored.
               If multiple_cartoons_displayed = False Then
                  frmMain.AutoRedraw = False
                  
                  CopyImageToForm Me, cartoon_x_position, cartoon_y_position, imaSource.width, imaSource.height, picSave
               
                  frmMain.AutoRedraw = True
               End If
            End If
            
            If cartoon_count Then
               ' Load the next cartoon image
               ' This is imported into a Image control which chnages size to accomodate the image - this is the simplest
               ' way to determine the image's dimensions regardless of the image file format.
               imaSource.Picture = LoadPicture(cartoon_archive_path & "\" & cartoon_list(next_cartoon_index))
               ' Resize the source PictureBox control, and copy the cartonn image into this control.
               picSource.width = imaSource.width
               picSource.height = imaSource.height
               picSource.Picture = imaSource.Picture
               ' Now resize the desktop area saving Picturebox control to be the same size - we'll use this to store what
               ' is on the desktop display before the cartoon is drawn.
               picSave.width = imaSource.width
               picSave.height = imaSource.height
               
               ' Determine the array index of the next cartoon to be displayed.
               If random_order = True Then
                  next_cartoon_index = next_cartoon_index + (Rnd * CSng(cartoon_count))
               Else
                  next_cartoon_index = next_cartoon_index + 1
               End If
               
               If next_cartoon_index >= cartoon_count Then
                  next_cartoon_index = next_cartoon_index - cartoon_count
               End If
               
               ' Generate randon position for the cartoon image, such that it is not 'clipped' off of the desktop display.
               cartoon_x_position = CInt(CDbl((Screen.width / Screen.TwipsPerPixelX) - picSource.width) * Rnd)
               cartoon_y_position = CInt(CDbl((Screen.height / Screen.TwipsPerPixelY) - picSource.height) * Rnd)
                              
               tmrDisplay.Tag = STATE_SHOWCARTOON
            Else
               ' If no cartoon images exist in the array, then reset this machine state and try again later.
               ' Possibly by then some website content may have been successfully loaded.
               timer_count = 0
            End If
         End If
         
      Case STATE_PREVIEW
         ' Screensaver Preview window state - used when panning the image showing all the Dilbert characters
         '                                    around in the preview window of the Display Properties dialog
         '                                    in Windows.
         
         ' Check that Display Properties dialog is still open (by checking for existance of supplied preview
         ' window handle. If not found, then terminate preview process.
         If HasPreviewWindowClosed Then
            tmrDisplay.Enabled = False
            Unload Me
         End If
         
         ' Thus is a polar equation that moves the x and y axis w.r.t a polar distance factor.
         ' If you don't have an understanding of polar mathematics, then don't mess with this!...
         dist = 25# - (Cos(preview_angle3) * 25#)
         
         picSplash.Top = (Cos(preview_angle1) * dist) - dist - 50
         picSplash.Left = (Sin(preview_angle2) * dist) - dist - 50
         picSplash.Visible = True
         preview_angle1 = preview_angle1 + PREVIEW_ANGLE_DIFF1
         preview_angle2 = preview_angle2 + PREVIEW_ANGLE_DIFF2
         preview_angle3 = preview_angle3 + PREVIEW_ANGLE_DIFF3
         
   End Select

End Sub

Private Sub DoUnload()
   
   ' Set signal to terminate any Timer activity as soon as possible.
   is_unloading = True
   Unload Me

End Sub

