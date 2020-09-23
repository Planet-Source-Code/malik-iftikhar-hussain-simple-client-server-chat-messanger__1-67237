Attribute VB_Name = "Module1"

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
      Public Type NOTIFYICONDATA
         cbSize As Long
         hwnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type

        Public Const NIM_ADD = &H0
   Public Const NIM_MODIFY = &H1
   Public Const NIM_DELETE = &H2

     Public Const WM_MOUSEMOVE = &H200

    Public Const NIF_MESSAGE = &H1
   Public Const NIF_ICON = &H2
   Public Const NIF_TIP = &H4

    
      'Declare the API function call.
 Public Const MAX_PATH = 260
Public MyName As String * 260
Dim nid As NOTIFYICONDATA





 Public Sub AddIcon()
         
         nid.cbSize = Len(nid)
         nid.hwnd = Form1.hwnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = Form1.Icon
         nid.szTip = Trim(MyName) & vbNullChar

         Shell_NotifyIcon NIM_ADD, nid
      End Sub


 
   Public Sub DeleteIcon()
          Shell_NotifyIcon NIM_DELETE, nid
    End Sub

Sub Main()
    Load Form1
    
End Sub
