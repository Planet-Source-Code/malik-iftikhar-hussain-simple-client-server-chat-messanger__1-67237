VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Box"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin MessageBox.xpButton CmdRply 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "Reply"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Form1.frx":1272
   End
   Begin VB.TextBox tXTINCMESSAGE 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   240
      Width           =   7575
   End
   Begin VB.CommandButton Cmdclose 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2640
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7815
      Begin MessageBox.xpButton Cmdsend 
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         TX              =   "Send"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Form1.frx":128E
      End
   End
   Begin MessageBox.xpButton Cmdok 
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Form1.frx":12AA
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   7920
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Iftikhar Malik,    Sialkot.     humsafar_ak@yahoo.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Iftikhar Malik,    Sialkot.     humsafar_ak@yahoo.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   2160
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This project first convert to exe and install it at Client's(Node) system, then add a shortcut at "START/PROGRAMS/STARTUP" menu.
' You can find your computer name popup on mouse over the icon at taskbar.
' type this name at the server project's computer name combo box.
' Any Question or comments please send a mail to cmshafi@yahoo.com
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*****************************NODE / CLIENT PROEJECT***************************
Private Sub Cmdok_Click()
Winsock1.SendData "OK"
End Sub


Private Sub Cmdsend_Click()
Winsock1.SendData txtSend.Text
 Cmdsend.Enabled = False
 txtSend.Text = ""
 txtSend.SetFocus
End Sub

Private Sub CmdRply_Click()
Cmdsend.Visible = True
txtSend.Visible = True
txtSend.SetFocus
CmdRply.Enabled = False
Cmdok.Enabled = False
End Sub

Private Sub Cmdclose_Click()
'Winsock1.SendData "Closed messagebox " & Time
'Winsock1.Close
'Me.Hide
End Sub

Private Sub Form_Load()
 GetComputerName MyName, MAX_PATH

AddIcon
'Me.Show 1
Me.Hide
Form1.Caption = " Message Box : " & Trim(MyName)

Winsock1.RemotePort = 1001
Winsock1.Bind 1002, Trim(MyName)
'winsock1.LocalHostName or
'winsock1.LocalIP can beuse to find computer name or ip address

Cmdsend.Enabled = False
Cmdsend.Visible = False
txtSend.Visible = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Winsock1.Close
End Sub

Private Sub txtSend_Change()
If txtSend.Text <> "" Then
Cmdsend.Enabled = True
Else
Cmdsend.Enabled = False
End If
End Sub



Private Sub winsock1_DataArrival _
(ByVal bytesTotal As Long)
Dim StrData As String
Dim dCOLOR As Long
Dim dBOLD As Boolean
Dim dSIZE As Single

On Error Resume Next

Winsock1.GetData StrData
Me.Show
Winsock1.SendData "Recieved Data"
'Shuttingdown the system
If StrData = "shutdowmsystem@gs6hge09hj4ljsg7??J((*&^@54wet6" Then
Winsock1.SendData "Shuting down"
tXTINCMESSAGE.Text = "This System will shutdown now"
ExitWindowsEx 1, 0

Else
'Closing messagebox
If StrData = "CloseMessagebox@7m3n4kscb82hddsj??!!$$$" Then
Winsock1.SendData "Message Box Closing"
Me.Hide
Else

'Set font
tXTINCMESSAGE.font = StrData
Winsock1.GetData dSIZE

tXTINCMESSAGE.font.Size = dSIZE
Winsock1.GetData dBOLD
tXTINCMESSAGE.font.Bold = dBOLD
Winsock1.GetData dCOLOR
tXTINCMESSAGE.ForeColor = dCOLOR
Winsock1.GetData StrData
tXTINCMESSAGE.Text = StrData

End If
End If

End Sub



      Private Sub Form_Terminate()
          DeleteIcon
      End Sub




