VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jarry Leo 聊天软件客户端"
   ClientHeight    =   6225
   ClientLeft      =   150
   ClientTop       =   600
   ClientWidth     =   7380
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   6225
   ScaleWidth      =   7380
   StartUpPosition =   2  '屏幕中心
   Begin 聊天室客户端.jcbutton Command1 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   5700
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16572874
      Caption         =   "发送(&S)"
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   5790
      Left            =   5040
      TabIndex        =   5
      Top             =   300
      Width           =   2115
   End
   Begin 聊天室客户端.jcbutton Command2 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   5700
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16572874
      Caption         =   "关闭(&C)"
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   3660
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4620
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3975
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "服务器IP"
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5700
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "通信状态："
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   4380
      Width           =   900
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ys As Integer
Dim zd As Boolean
Public WithEvents winsock1 As CSocketMaster
Attribute winsock1.VB_VarHelpID = -1

Private Sub Command1_Click() '发送信息
If winsock1.State = 7 Then '检测连接状态
    winsock1.SendData JiaMi("Spek" & Text2, "")
    Text2 = ""
Else
    MsgBox "当前处于离线状态，无法发送消息！", vbOKOnly + 64, "提示"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Sub Form_Load() '连接服务器
Set winsock1 = New CSocketMaster
winsock1.RemoteHost = SVR_IP
winsock1.RemotePort = SVR_PORT
Label2 = "当前服务器IP：" & SVR_IP
winsock1.Connect
zd = True
Me.Caption = Me.Caption + " [" & USER_NAME & "]"
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_DblClick() '双击私聊
If winsock1.State = 7 And List1.Text <> USER_NAME Then
    Dim oneQ As New ONE
    oneQ.Show
    oneQ.Caption = "与" & List1.Text & "的对话！"
End If
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    Command1_Click
End If
End Sub

Private Sub Timer1_Timer() '连接状态管理
Dim st(9) As String, t As Integer

st(0) = "等待重新连接...[" & Int(ys / 10) & "秒]"
st(1) = "打开"
st(2) = "等待连接"
st(3) = "连接挂起"
st(4) = "解析域名"
st(5) = "识别主机"
st(6) = "正在连接..."
st(7) = "已连接"
st(8) = "连接断开"
st(9) = "连接错误"
Label1 = "通信状态：" & st(winsock1.State)
Label1.ForeColor = vbBlue

If winsock1.State = 7 Then Label1.ForeColor = vbGreen
If winsock1.State > 7 Then
Label1.ForeColor = vbRed
If zd = True Then ys = ys + 2
End If
If winsock1.State = 0 And ys > 0 Then Label1.ForeColor = vbBlack: ys = ys - 1
If ys > 60 Then winsock1.CloseSck
If ys = 1 Then
    If winsock1.State = 0 Then
    winsock1.Connect
    End If
End If
    

End Sub

Private Sub Winsock1_Connect() '当连接成功时给服务器发送信息自报家门
winsock1.SendData JiaMi("Name" & USER_NAME, "")
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long) '接收数据
Dim tempS As String
Dim ii As Integer
Dim frm As Form
Dim cc As String
winsock1.GetData tempS  '读取数据
tempS = JieMi(tempS, "")
Select Case Left(tempS, 4)
Case "List" '接收列表
    List1.Clear
    ll = Split(Right(tempS, Len(tempS) - 4), "List")
        For Each H In ll
            List1.AddItem H
        Next H
Case "Spek" '接受对话
    If Text1 = "" Then
        Text1 = Text1 & Right(tempS, Len(tempS) - 4) & vbCrLf
        Else
        Text1 = Text1 & vbCrLf & Right(tempS, Len(tempS) - 4) & vbCrLf
    End If
Case "Scut" '接收系统信息
    Text1 = Text1 & vbCrLf & Right(tempS, Len(tempS) - 4)
    zd = False
End Select

If Left(tempS, 2) = "To" Then  '接收私聊
    ii = Val(Mid(tempS, 3, 2))
    cc = "与" & List1.List(ii) & "的对话！"
    
    For Each frm In Forms
        If frm.Caption = cc Then
            frm.Text1 = frm.Text1 & List1.List(ii) & " (" & Now & ") " & "说：" & vbCrLf & Right(tempS, Len(tempS) - 4) & vbCrLf & vbCrLf
            Exit Sub
        End If
    Next
    Dim oneW As New ONE
    oneW.Show
    oneW.Caption = "与" & List1.List(ii) & "的对话！"
    oneW.Text1 = List1.List(ii) & " (" & Now & ") " & "说：" & vbCrLf & Right(tempS, Len(tempS) - 4) & vbCrLf & vbCrLf
End If
End Sub
