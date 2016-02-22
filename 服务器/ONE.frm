VERSION 5.00
Begin VB.Form ONE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "与某某的对话"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6930
   Icon            =   "ONE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6930
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "关闭(&C)"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "发送(&S)"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4200
      Width           =   6495
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
      Width           =   6495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现实地址："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对方的IP地址："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   5700
      Width           =   1260
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

For yy = 1 To Form1.List1.ListCount - 1
    If Form1.List1.List(yy) = mubiao Then
        Form1.Connect1(yy).SendData JiaMi("To" & Format(0, "00") & Text2, "")
        Text1 = Text1 & "admin (" & Now & ") " & "说：" & vbCrLf & Text2 & vbCrLf & vbCrLf
        Text2 = ""
    End If
Next yy
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim m As String
Dim mubiao As String
Dim yy As Integer
mu = Me.Caption
mubiao = Mid(mu, 2, Len(mu) - 5)

Dim c As String
Dim a, b As Long
For yy = 1 To Form1.List1.ListCount - 1
If Form1.List1.List(yy) = mubiao Then
    c = Form1.Connect1(yy).RemoteHostIP
    Label1 = "对方的IP地址：" & c
    Label2 = IPadress(c) '对方的现实地址

End If
Next yy
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
Function IPadress(ip As String) As String '获取IP地址的现实地址
On Error GoTo err
Dim a As Integer, b As Integer
Dim j As String
j = HtmlStr("http://www.ip138.com/ips138.asp?ip=" & ip & "&action=2")
a = InStr(1, j, "<li>")
b = InStr(a, j, "</li>")
IPadress = Mid(j, a + 10, b - a - 10)
Exit Function
err:
IPadress = ""
End Function
Function HtmlStr$(Url$) '提取网页源码函数
Dim XmlHttp
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
    XmlHttp.Open "GET", Url, False
    XmlHttp.send
If XmlHttp.ReadyState = 4 Then HtmlStr = StrConv(XmlHttp.ResponseBody, vbUnicode)
Set XmlHttp = Nothing
End Function
