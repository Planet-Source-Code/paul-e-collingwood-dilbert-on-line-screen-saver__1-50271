Attribute VB_Name = "modDownload"
Option Explicit

Private Const INTERNET_OPEN_TYPE_DIRECT As Long = 1

Private Const INTERNET_SERVICE_HTTP  As Long = 3

Private Const INTERNET_FLAG_NO_COOKIES As Long = &H80000
Private Const INTERNET_FLAG_NO_CACHE_WRITE As Long = &H4000000

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" ( _
   ByVal lpszAgent$, _
   ByVal dwAccessType&, _
   ByVal lpszProxyName$, _
   ByVal lpszProxyBypass$, _
   ByVal dwFlags&) As Long

Private Declare Function InternetCloseHandle Lib "wininet" ( _
   ByVal hEnumHandle&) As Long
      
Private Declare Function InternetConnect Lib "wininet" Alias "InternetConnectA" ( _
   ByVal internet_handle&, _
   ByVal lpszServerName$, _
   ByVal nServerPort&, _
   ByVal lpszUserName$, _
   ByVal lpszPassword$, _
   ByVal dwService&, _
   ByVal dwFlags&, _
   ByVal context_word&) As Long

Private Declare Function HttpOpenRequest Lib "wininet" Alias "HttpOpenRequestA" ( _
    ByVal hHttpSession&, _
    ByVal lpszVerb$, _
    ByVal lpszObjectName$, _
    ByVal lpszVersion$, _
    ByVal lpszReferer$, _
    ByVal lpszAcceptTypes$, _
    ByVal dwFlags&, _
    ByVal dwContext&) As Long

Private Declare Function HttpSendRequest Lib "wininet" Alias "HttpSendRequestA" ( _
    ByVal hHttpRequest&, _
    ByVal lpszHeaders$, _
    ByVal dwHeadersLength&, _
    ByVal lpOptional$, _
    ByVal dwOptionalLength&) As Boolean

Private Declare Function InternetQueryDataAvailable Lib "wininet" ( _
    ByVal hFile&, _
    ByRef lpdwNumberOfBytesAvailable&, _
    ByVal dwFlags&, _
    ByVal dwContext&) As Boolean

Private Declare Function InternetReadFile Lib "wininet" ( _
    ByVal hFile&, _
    ByVal lpBuffer$, _
    ByVal dwNumberOfBytesToRead&, _
    ByRef lpNumberOfBytesRead&) As Boolean


Public Function DownloadURLToString(ByVal server_url$, ByVal path_url$, ByRef return_string$) As Boolean
   Dim open_handle&, connect_handle&, request_handle&, result As Boolean, bytes_read&, total_string$
   Dim string_buffer As String * 1024
   
   On Error Resume Next
   
   DownloadURLToString = False
   
   open_handle = InternetOpen(App.Title, _
                              INTERNET_OPEN_TYPE_DIRECT, _
                              vbNullString, vbNullString, 0)
   
   If open_handle Then
   
      connect_handle = InternetConnect(open_handle, _
                                       server_url, _
                                       80, _
                                       "", "", _
                                       INTERNET_SERVICE_HTTP, _
                                       0, 0)
      If connect_handle Then
      
         request_handle = HttpOpenRequest(connect_handle, _
                                          "GET", _
                                          path_url, _
                                          "HTTP/1.0", vbNullString, vbNullString, _
                                          INTERNET_FLAG_NO_COOKIES Or INTERNET_FLAG_NO_CACHE_WRITE, _
                                          0)

         If request_handle Then

            result = HttpSendRequest(request_handle, vbNullString, 0, vbNullString, 0)
   
            If result Then
               
               Do
                  DoEvents
                  result = InternetReadFile(request_handle, string_buffer, Len(string_buffer), bytes_read)
                  
                  If result Then
                     If bytes_read > 0 Then
                        total_string = total_string & Left$(string_buffer, bytes_read)
                     End If
                  Else
                     Exit Do
                  End If
                  
               Loop While bytes_read > 0
   
               If result Then
                  return_string = total_string
               
                  DownloadURLToString = True
               End If
               
            End If
   
            InternetCloseHandle request_handle
         
         End If
   
         InternetCloseHandle connect_handle
         
      End If

      InternetCloseHandle open_handle

   End If
   
End Function

Public Function DownloadURLToFile(ByVal server_url$, ByVal path_url$, ByVal destination_file_path$) As Boolean
   Dim file_number%, string_buffer$

   On Error Resume Next
   
   DownloadURLToFile = False
   
   Err.Clear
   file_number = FreeFile
   If Err.Number = 0 Then
   
      Open destination_file_path For Binary Access Write As #file_number
      If Err.Number = 0 Then
      
         If DownloadURLToString(server_url, path_url, string_buffer) Then
            
            Err.Clear
            Put #file_number, , string_buffer
            If Err.Number = 0 Then
            
               DownloadURLToFile = True
            
            End If
            
         End If
         
         Close #file_number
         
      End If
   End If
End Function

