VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "SMSICQ"
   ClientHeight    =   6585
   ClientLeft      =   4935
   ClientTop       =   2145
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   5025
   Begin VB.Frame statuspanel 
      Caption         =   " Status "
      Height          =   630
      Left            =   15
      TabIndex        =   10
      Top             =   5520
      Width           =   4950
      Begin VB.Label status 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   255
         Width           =   4680
      End
   End
   Begin VB.TextBox number 
      Height          =   300
      Left            =   3120
      TabIndex        =   3
      Top             =   2415
      Width           =   1800
   End
   Begin VB.TextBox prefix 
      Height          =   300
      Left            =   1290
      TabIndex        =   2
      Top             =   2190
      Width           =   630
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4365
      Top             =   1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   345
      Left            =   3675
      TabIndex        =   5
      Top             =   5085
      Width           =   1260
   End
   Begin VB.TextBox msg 
      Height          =   1860
      Left            =   60
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3135
      Width           =   4875
   End
   Begin VB.TextBox pass 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   1515
      Width           =   1920
   End
   Begin VB.TextBox user 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   1065
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   5010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ex. Portugal Code = 351"
      Height          =   195
      Index           =   4
      Left            =   165
      TabIndex        =   15
      Top             =   2610
      Width           =   1725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   5040
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   2
      X1              =   0
      X2              =   5040
      Y1              =   2925
      Y2              =   2925
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   5040
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   0
      X1              =   0
      X2              =   5040
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Label email 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bruno@escripovoa.pt"
      Height          =   195
      Left            =   3390
      TabIndex        =   14
      Top             =   6225
      Width           =   1530
   End
   Begin VB.Label copyrigth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyrigth 2002 - Bruno Coelho"
      Height          =   195
      Left            =   75
      TabIndex        =   13
      Top             =   6210
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "You have :"
      Height          =   195
      Left            =   45
      TabIndex        =   12
      Top             =   5115
      Width           =   780
   End
   Begin VB.Label words 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   885
      TabIndex        =   11
      Top             =   5100
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number :"
      Height          =   195
      Index           =   3
      Left            =   3315
      TabIndex        =   9
      Top             =   2145
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country Code :"
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   8
      Top             =   2220
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   7
      Top             =   1575
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icq # :"
      Height          =   195
      Index           =   0
      Left            =   510
      TabIndex        =   6
      Top             =   1095
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Dim ErrMsg As String
status.Caption = "Opening registry page and say you are online..."
'opens the registry page and say your are online
retval = Inet1.OpenURL("http://web.icq.com/karma/dologin/1,,,00.html?uService=1&uLogin=" + user.Text + "&uPassword=" + pass.Text)
status.Caption = "Sending the message to the phone number"
'send the message to the phone number you want
retval = Inet1.OpenURL("https://web.icq.com/secure/sms/send_history/1,,,00.html?country=" + prefix.Text + "&carrier=aaa&tophone=" + number.Text + "&y=15&prefix=%2B" + prefix.Text + "&uSend=1&charcount=150&msg=" + msg.Text)
'check if message was sent successfully
If InStr(1, retval, "The SMS message could not be sent.") > 0 Then
    ErrMsg = "The SMS message could not be sent." + vbCrLf
    'check if the error is because you reach the message-sending limit
    If InStr(1, retval, "You have reached your daily SMS message-sending limit") > 0 Then
        ErrMsg = ErrMsg + "You have reached your daily SMS message-sending limit." + vbCrLf
        ErrMsg = ErrMsg + "Please try again later."
    End If
    junk = MsgBox(ErrMsg, vbInformation + vbOKOnly, "SmsIcq")
Else
    junk = MsgBox("Your message was sent with success !", vbInformation + vbOKOnly, "SmsIcq")
End If
status.Caption = ""
End Sub



Private Sub email_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
email.ForeColor = QBColor(9)
email.FontUnderline = True
End Sub

Private Sub email_Click()
isTemp = "mailto:bruno@escripovoa.pt"
lRet = ShellExecute(hWnd, "open", isTemp, vbNull, vbNull, 1)

End Sub

Private Sub Form_Load()
words.Caption = 150
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
email.ForeColor = QBColor(0)
email.FontUnderline = False

End Sub

Private Sub Label3_Click()

End Sub

Private Sub msg_Change()
words.Caption = 150 - Len(msg)
End Sub
