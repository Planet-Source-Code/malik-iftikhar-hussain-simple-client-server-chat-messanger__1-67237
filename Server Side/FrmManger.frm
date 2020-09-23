VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmManger 
   Caption         =   "Net Work Manager"
   ClientHeight    =   5040
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7800
   Icon            =   "FrmManger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear Message"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox CmboNode 
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton CmdShutCli 
      Caption         =   "Shutdown System"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton CmdCloseCli 
      Caption         =   "Close Message box"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   720
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "Send Message"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox TxtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Frame Framex 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   7455
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         Caption         =   "Message to Send"
         Height          =   1215
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   6975
      End
      Begin VB.ListBox ListReply 
         Height          =   840
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   6735
      End
      Begin VB.CommandButton CmdFont 
         Caption         =   "Change Font"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Caption         =   "Read Message"
         Height          =   1215
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   6975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C000&
         BackStyle       =   1  'Opaque
         Height          =   3375
         Left            =   120
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Iftikhar Malik,     Sialkot.  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      DrawMode        =   9  'Not Mask Pen
      Index           =   0
      X1              =   120
      X2              =   7560
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address /  Computer Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2685
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   7560
      Y1              =   4320
      Y2              =   4320
   End
End
Attribute VB_Name = "FrmManger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const MAX_PATH = 260
Dim MyName As String * 260
Dim LastData As Long
 Private Type record1
    mESSAGE As String * 200
    mESFONT As String * 15
    mESfONTsIZE As Integer
    mESfONTbOLD As Boolean
    mESfOREcOLOR As Long
End Type
Private Type Record2
    SistMe As String * 30
    SendMe As Boolean
    LockMe As Boolean
End Type
Private Sub cONnectNod()
With Winsock1
.RemoteHost = CmboNode.Text
.RemotePort = 1002   ' Port to connect to.
.Bind 1001                ' Bind to the local port.

End With
End Sub






Private Sub CmdFont_Click()
' Set Cancel to True
  CommonDialog1.CancelError = True
  On Error Resume Next
  ' Set the Flags property
  CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
  ' Display the Font dialog box
  CommonDialog1.FontName = TxtMessage.FontName
  CommonDialog1.FontBold = TxtMessage.FontBold
  CommonDialog1.FontSize = TxtMessage.FontSize
  CommonDialog1.Color = TxtMessage.ForeColor
  CommonDialog1.ShowFont
  TxtMessage.Font.Name = CommonDialog1.FontName
 TxtMessage.FontSize = CommonDialog1.FontSize
 TxtMessage.Font.Bold = CommonDialog1.FontBold
'  Text1.Font.Italic = CommonDialog1.FontItalic
'  Text1.Font.Underline = CommonDialog1.FontUnderline
'  Text1.FontStrikethru = CommonDialog1.FontStrikethru
 TxtMessage.ForeColor = CommonDialog1.Color

 TxtMessage.SetFocus


End Sub

Private Sub CmdRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
CmdRefresh.BackColor = &H800000
End Sub

Private Sub CmdShutCli_Click()
Winsock1.SendData "shutdowmsystem@gs6hge09hj4ljsg7??J((*&^@54wet6"
'Winsock1.Close
End Sub







Private Sub Label2_Click(Index As Integer)
frmAbout.Show
End Sub



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
'On Error GoTo ERRORHANDLER
Winsock1.GetData strData

ListReply.AddItem strData & vbTab & Time
 If strData = "Shuting down" Or strData = "Message Box Closing" Then
  Winsock1.Close
CmdCloseCli.Enabled = False
 CmdShutCli.Enabled = False
 
Else
CmdCloseCli.Enabled = True
 CmdShutCli.Enabled = True


 End If
'ERRORHANDLER:
'             MsgBox Err.Description, vbInformation, "Remote System"
'             End
End Sub



Private Sub CmboNode_Change()
Framex.Caption = CmboNode.Text
Winsock1.Close
cONnectNod
End Sub

Private Sub CmboNode_Click()
Winsock1.Close
cONnectNod
End Sub

Private Sub CmboNode_DblClick()
Winsock1.Close
cONnectNod
End Sub

Private Sub CmdClear_Click()
TxtMessage.Text = ""
TxtMessage.SetFocus
CmdClear.Enabled = False
End Sub

Private Sub CmdCloseCli_Click()
Winsock1.SendData "CloseMessagebox@7m3n4kscb82hddsj??!!$$$"

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub



Private Sub CmdSend_Click()
Winsock1.SendData TxtMessage.Font
Winsock1.SendData TxtMessage.FontSize
Winsock1.SendData TxtMessage.Font.Bold
Winsock1.SendData TxtMessage.ForeColor
Winsock1.SendData TxtMessage.Text
TxtMessage = ""
TxtMessage.SetFocus
CmdSend.Enabled = False

End Sub



Private Sub Form_Load()
GetComputerName MyName, MAX_PATH
If CmboNode.Text = "" Then CmboNode.Text = Trim(MyName)
Framex.Caption = CmboNode.Text
CmdSend.Enabled = False
CmdClear.Enabled = False
End Sub



Private Sub TxtMessage_Change()

If TxtMessage.Text <> "" Then

CmdSend.Enabled = True
CmdClear.Enabled = True
Else

CmdSend.Enabled = False
CmdClear.Enabled = False
End If
End Sub



