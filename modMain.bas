Attribute VB_Name = "modMain"
Option Explicit

' Windows API calls (and support defintions) used by this screensaver engine
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1

Private Const WS_CHILD = &H40000000
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest

Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const NUMCOLORS = 24         '  Number of colors the device supports

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
   ByVal hwnd&, _
   ByVal nIndex&) As Long

Private Declare Function GetClientRect Lib "user32" ( _
   ByVal hwnd&, _
   lpRect As RECT) As Long
   
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
   ByVal hwnd&, _
   ByVal nIndex&, _
   ByVal dwNewLong&) As Long
   
Private Declare Function SetParent Lib "user32" ( _
   ByVal hWndChild&, _
   ByVal hWndNewParent&) As Long
   
Private Declare Function SetWindowPos Lib "user32" ( _
   ByVal hwnd&, _
   ByVal hWndInsertAfter&, _
   ByVal X&, _
   ByVal Y&, _
   ByVal cx&, _
   ByVal cy&, _
   ByVal wFlags&) As Long
   
Private Declare Function ShowCursor Lib "user32" ( _
   ByVal bShow&) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
   ByVal lpClassName$, _
   ByVal lpWindowName$) As Long

Private Declare Function GetDesktopWindow Lib "user32" ( _
   ) As Long

Private Declare Function GetDC Lib "user32" ( _
   ByVal hwnd&) As Long

Private Declare Function BitBlt Lib "gdi32" ( _
   ByVal hDestDC&, _
   ByVal X&, _
   ByVal Y&, _
   ByVal nWidth&, _
   ByVal nHeight&, _
   ByVal hSrcDC&, _
   ByVal xSrc&, _
   ByVal ySrc&, _
   ByVal dwRop&) As Long

Private Declare Function IsWindow Lib "user32" ( _
   ByVal hwnd&) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
   ByVal hdc&, _
   ByVal nIndex&) As Long

Private Declare Function StretchBlt Lib "gdi32" ( _
   ByVal hdc&, _
   ByVal X&, _
   ByVal Y&, _
   ByVal nWidth&, _
   ByVal nHeight&, _
   ByVal hSrcDC&, _
   ByVal xSrc&, _
   ByVal ySrc&, _
   ByVal nSrcWidth&, _
   ByVal nSrcHeight&, _
   ByVal dwRop&) As Long

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" ( _
   ByVal hwnd&, _
   ByVal lpOperation$, _
   ByVal lpFile$, _
   ByVal lpParameters$, _
   ByVal lpDirectory$, _
   ByVal nShowCmd&) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
   ByVal nBufferLength&, _
   ByVal lpBuffer$) As Long


Private Declare Function GetWindowRect Lib "user32" ( _
   ByVal hwnd&, _
   lpRect As RECT) As Long
      
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
   ByVal lpLibFileName$) As Long

Private Declare Function FreeLibrary Lib "kernel32" ( _
   ByVal hLibModule&) As Long
   
Private Declare Function GetProcAddress Lib "kernel32" ( _
   ByVal hModule&, _
   ByVal lpProcName$) As Long
   
Private Declare Function AlphaBlend Lib "msimg32" ( _
   ByVal hDestDC As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
    
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   Destination As Any, _
   Source As Any, _
   ByVal Length&)
    
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

' Internal variables

Private Enum RunningMode
   Configuration = 1
   PREVIEW = 2
   SCREENSAVER = 3
End Enum

Private Const APP_NAME = "Dilbert Screen Saver (active)"

Private preview_window_handle&
Private display_properties_window_handle&
Private command_line_arguments$
Private running_mode As RunningMode

Public application_name$

Private Sub DetectPreviousInstance()
   
   On Error Resume Next
   
   ' Uses fail-safe mechanism to detect a previous instance of this application (running in SCREENSAVER mode)
   ' If NO instance found, then it's safe to execute this application, so returns from this subroutine

    If App.PrevInstance Then
      If FindWindow(vbNullString, APP_NAME) Then
         End
      End If
   End If
   
End Sub

Public Sub Main()
   Dim preview_rect As RECT, window_style&, desktop_handle&, desktop_device_context&

   On Error Resume Next

   application_name = App.Title
   
   desktop_handle = GetDesktopWindow
   desktop_device_context = GetDC(desktop_handle)
   
    ' Get the command line arguments.
   command_line_arguments = LCase(Trim(Command$))
   'MsgBox command_line_arguments
   
   ' Check the first 2 characters to determine the requested mode of operation
   Select Case Left$(command_line_arguments, 2)
      Case ""
         Err.Clear
         Debug.Print 1 / 0
         If Err.Number Then
            ' Contrived situation used only when in the VB IDE -
            ' displays the debug dialog so that all the screensaver functionalty can be tested
            ' functionality can be tested.
            running_mode = SCREENSAVER
            frmDebug.Show 1
         Else
            ' Request for configuration dialog.
            running_mode = Configuration
         End If
      
      Case "/c"
         ' Request for configuration dialog.
         running_mode = Configuration
        
      Case "/s"
         ' Request for standard operation as a screen saver.
         running_mode = SCREENSAVER
            
      Case "/p"
         ' Request for preview mode.
         running_mode = PREVIEW
            
      Case Else
         ' We shouldn't get any other type of command line, so just run as screensaver.
         running_mode = SCREENSAVER
   End Select

   Select Case running_mode
      Case Configuration
         ' Determine the handle of the calling 'Display Properties' dialog (if supplied)
         ' This is used to periodlically check that this parent window is still open - if closed
         ' then the sceensaver setting dialog will also close.
         If command_line_arguments = "" Then
            display_properties_window_handle = 0
         Else
            display_properties_window_handle = ExtractWindowHandle()
         End If
                        
         frmSettings.Show
                
      Case SCREENSAVER
            
         DetectPreviousInstance
            
         frmMain.Show
         ' Set the screensaver form's caption so that subsequent instances of this application will
         ' detect this instance in DetectPreviousInstance()
         frmMain.Caption = APP_NAME
            
         ' Check to see if the cursor is to be hidden while the screensaver is operative.
         Select Case Trim(UCase(CStr(GetSetting(App.Title, "Configuration", "HideCursor", "NO"))))
            Case "YES", "Y", "1"
               ShowCursor False
         End Select
            
      Case PREVIEW
         'The screensaver form is loaded, but not shown.
         Load frmMain
            
         ' It is now prepared to display in the preview porthole on the 'Display Properties' dialog
         frmMain.Caption = "Preview"

         ' Get the current window style...
         window_style = GetWindowLong(frmMain.hwnd, GWL_STYLE)

         ' ...alter it to denote it is a child window...
         window_style = (window_style Or WS_CHILD)

         ' ..and make the change permanent.
         SetWindowLong frmMain.hwnd, GWL_STYLE, window_style

         ' Determine the handle of the calling 'Display Properties' dialog's preview window.
         preview_window_handle = ExtractWindowHandle()

         ' Now we set our screensaver form's parent to be this preview window.
         ' Our form will now appear 'inside' the the preview area!
         SetParent frmMain.hwnd, preview_window_handle
         SetWindowLong frmMain.hwnd, GWL_HWNDPARENT, preview_window_handle

         ' Lastly, we need to resize our screensaver form to fit entirely into the parent preview window,
         ' so firstly we determine the preview window's dimensions...
         GetClientRect preview_window_handle, preview_rect
            
         ' ...and then we resize our form and show it simultaneously.
         SetWindowPos frmMain.hwnd, HWND_TOP, _
                        0&, 0&, preview_rect.Right, preview_rect.Bottom, _
                        SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    End Select
    
End Sub

Private Function ExtractWindowHandle() As Long
   Dim index%, handle_character$

   On Error Resume Next

   For index = Len(command_line_arguments) To 1 Step -1
      handle_character = Mid(command_line_arguments, index, 1)
      If handle_character < "0" Or handle_character > "9" Then
         Exit For
      End If
    Next

    ExtractWindowHandle = CLng(Mid(command_line_arguments, index + 1))
    
End Function

Public Sub SetWindowTopmost(the_form As Form)
   On Error Resume Next
   
   SetWindowPos the_form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
   
End Sub

Public Sub CopyDesktopToForm(the_form As Form)
   Dim desktop_window_handle&, desktop_device_context&

   On Error Resume Next
   
   desktop_window_handle = GetDesktopWindow
   desktop_device_context = GetDC(desktop_window_handle)

   BitBlt the_form.hdc, 0, 0, Screen.width, Screen.height, desktop_device_context, 0, 0, SRCCOPY

End Sub

Public Sub UnHideCursor()
   
   ShowCursor True

End Sub

Public Sub CopyImageToForm(the_form As Form, ByVal X&, ByVal Y&, ByVal width&, ByVal height&, source_control As Control)
   
   BitBlt the_form.hdc, X, Y, width, height, source_control.hdc, 0, 0, SRCCOPY
   
End Sub

Public Sub CopyImageFromForm(the_form As Form, ByVal X&, ByVal Y&, ByVal width&, ByVal height&, dest_control As Control)
   
   BitBlt dest_control.hdc, 0, 0, width, height, the_form.hdc, X, Y, SRCCOPY

End Sub

Public Sub AndImageToImage(dest_control As Control, source_control As Control, fade_control As Control)
   
   BitBlt dest_control.hdc, 0, 0, dest_control.width, dest_control.height, source_control.hdc, 0, 0, SRCCOPY
   BitBlt dest_control.hdc, 0, 0, dest_control.width, source_control.height, fade_control.hdc, 0, 0, SRCAND

End Sub

Public Sub AndImageToForm(the_form As Form, source_control As Control)
   
   BitBlt the_form.hdc, 0, 0, source_control.width, source_control.height, source_control.hdc, 0, 0, SRCAND

End Sub

Public Sub AlphaBlendImageToForm(the_form As Form, source_control As Control, ByVal fade_factor%)
    Dim Blend As BLENDFUNCTION, BlendLng As Long
   
    Blend.SourceConstantAlpha = fade_factor
    
    CopyMemory BlendLng, Blend, 4
    
    AlphaBlend the_form.hdc, 0, 0, source_control.width, source_control.height, _
               source_control.hdc, 0, 0, source_control.width, source_control.height, BlendLng

End Sub

Public Sub AlphaBlendImageToImage(dest_control As Control, base_control As Control, source_control As Control, temp_control As Control, ByVal fade_factor%)
    Dim Blend As BLENDFUNCTION, BlendLng As Long
   
    BitBlt temp_control.hdc, 0, 0, source_control.width, source_control.height, base_control.hdc, 0, 0, SRCCOPY
   
    Blend.SourceConstantAlpha = fade_factor
    
    CopyMemory BlendLng, Blend, 4
    
    AlphaBlend temp_control.hdc, 0, 0, source_control.width, source_control.height, _
               source_control.hdc, 0, 0, source_control.width, source_control.height, BlendLng

    BitBlt dest_control.hdc, 0, 0, source_control.width, source_control.height, temp_control.hdc, 0, 0, SRCCOPY

End Sub

Public Sub TileImageToImage(dest_control As Control, source_control As Control)
   Dim X&, Y&
   
   For X = 0 To dest_control.ScaleWidth Step source_control.ScaleWidth
      For Y = 0 To dest_control.ScaleHeight Step source_control.ScaleHeight
         BitBlt dest_control.hdc, X, Y, source_control.ScaleWidth, source_control.ScaleHeight, source_control.hdc, 0, 0, SRCCOPY
      Next Y
   Next X
End Sub


Public Function HasPreviewWindowClosed() As Boolean
   
   On Error Resume Next
   
   HasPreviewWindowClosed = True
   
   If IsWindow(preview_window_handle) Then
      HasPreviewWindowClosed = False
   End If

End Function

Public Function IsConfigParentWindowDefined() As Boolean
   
   If display_properties_window_handle Then
      IsConfigParentWindowDefined = True
   Else
      IsConfigParentWindowDefined = False
   End If

End Function

Public Function HasConfigParentWindowClosed() As Boolean
   
   On Error Resume Next
   
   HasConfigParentWindowClosed = True
   
   If IsWindow(display_properties_window_handle) Then
      HasConfigParentWindowClosed = False
   End If

End Function

Public Sub ShellURL(ByVal url$)
  
  On Error Resume Next
  
  Call ShellExecute(0&, vbNullString, url, vbNullString, vbNullString, vbNormalFocus)

End Sub

Public Function RunningAsScreensaver() As Boolean

   If running_mode = SCREENSAVER Then
      RunningAsScreensaver = True
   Else
      RunningAsScreensaver = False
   End If

End Function

Public Sub DebugSetScreensaverMode()
   
   running_mode = SCREENSAVER

End Sub

Public Sub DebugSetPreviewMode(ByVal args$)
   
   running_mode = PREVIEW
   command_line_arguments = Trim(args)

End Sub

Public Sub DebugSetConfigurationMode()
   
   running_mode = Configuration

End Sub

Public Function GetCartoonFileArchivePath() As String
   Dim temporary_path_length&, temporary_path$
   Dim temporary_path_buffer As String * 1000
   
   On Error Resume Next
   
   temporary_path_length = GetTempPath(1000, temporary_path_buffer)
   temporary_path = Left$(temporary_path_buffer, temporary_path_length)
   If Right$(temporary_path, 1) <> "\" Then
      temporary_path = temporary_path & "\"
   End If
   GetCartoonFileArchivePath = temporary_path & "Dilbert Cartoons"

End Function

Public Sub CentreSettingsForm(the_form As Form)
   Dim parent_rect As RECT, pos&
   
   On Error GoTo ERROR_EXIT
   
   If display_properties_window_handle <> 0 Then
      If GetWindowRect(display_properties_window_handle, parent_rect) Then
         pos = ((parent_rect.Left + parent_rect.Right) \ 2&) * Screen.TwipsPerPixelX
         pos = pos - (the_form.width \ 2&)
         If pos < 0 Then
            pos = 0
         ElseIf (pos + the_form.width) > Screen.width Then
            pos = Screen.width - the_form.width
         End If
         the_form.Left = pos
      
         pos = ((parent_rect.Top + parent_rect.Bottom) \ 2&) * Screen.TwipsPerPixelY
         pos = pos - (the_form.height \ 2&)
         If pos < 0 Then
            pos = 0
         ElseIf (pos + the_form.height) > Screen.height Then
            pos = Screen.height - the_form.height
         End If
         the_form.Top = pos
      End If
   End If
   
ERROR_EXIT:
End Sub

Public Function UseSimpleFading() As Boolean
   Dim desktop_handle&, desktop_device_context&, library_handle&, process_address&
   
   On Error Resume Next
   
   UseSimpleFading = True
   
   ' Check to see if the Alpha-blending API functionality is available to this version
   ' of windows
   library_handle = LoadLibrary("msimg32")
   process_address = GetProcAddress(library_handle, "AlphaBlend")
   FreeLibrary library_handle
   
   ' If relevant DLL procedure is not found, then exit.
   If process_address = 0 Then
      Exit Function
   End If
   
   ' Now check that we are running in a higher order screen colour depth.
   ' i.e either 16, 24, or 32-bit pallettes - if so we can use the alphablend
   ' method to acieve fading effects
   desktop_handle = GetDesktopWindow
   desktop_device_context = GetDC(desktop_handle)
   
   If GetDeviceCaps(desktop_device_context, BITSPIXEL) > 8 Then
      UseSimpleFading = False
   End If

End Function
