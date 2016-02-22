VERSION 5.00
Begin VB.Form ONE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "与某某的对话"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   Icon            =   "ONE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ONE.frx":1082
   ScaleHeight     =   6315
   ScaleWidth      =   6360
   StartUpPosition =   3  '窗口缺省
   Begin 聊天室客户端.jcbutton Command1 
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   5640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ButtonStyle     =   5
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   11169024
      Caption         =   "发送(&S)"
      ForeColor       =   16777215
   End
   Begin 聊天室客户端.jcbutton Command2 
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      ButtonStyle     =   5
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   11169024
      Caption         =   "关闭(&C)"
      ForeColor       =   16777215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4200
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "ONE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim m As String
Dim mubiao As String
Dim yy As Integer
mu = Me.Caption
mubiao = Mid(mu, 2, Len(mu) - 5)

For yy = 0 To FormMain.List1.ListCount - 1
If FormMain.List1.List(yy) = mubiao Then
    FormMain.winsock1.SendData JiaMi("To" & Format(yy, "00") & Text2, "")
    Text1 = Text1 & USER_NAME & " (" & Now & ") " & "说：" & vbCrLf & Text2 & vbCrLf & vbCrLf
    Text2 = ""
End If
Next yy
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1)
Beep
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Command1_Click
End If
End Sub
