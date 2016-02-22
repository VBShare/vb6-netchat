VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "聊天室服务器管理程序"
   ClientHeight    =   7425
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9405
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9405
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "服务器状态"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   6240
      Width           =   3015
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "在 线 数："
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "连 接 数："
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行状态："
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5520
      Top             =   3000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "发送公屏消息"
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5040
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   4575
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock Connect1 
      Index           =   0
      Left            =   6000
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock listen1 
      Left            =   6480
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "在线列表"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command2 
         Caption         =   "踢出房间"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "全选/全不选"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   5520
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   5280
         ItemData        =   "Form1.frx":1082
         Left            =   120
         List            =   "Form1.frx":1089
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "大厅"
      Height          =   7095
      Left            =   3240
      TabIndex        =   11
      Top             =   120
      Width           =   6015
   End
   Begin VB.Menu Control 
      Caption         =   "菜单"
      Begin VB.Menu Start_Server 
         Caption         =   "开始服务"
      End
      Begin VB.Menu Stop_Server 
         Caption         =   "停止服务"
      End
      Begin VB.Menu N1 
         Caption         =   "-"
      End
      Begin VB.Menu S_run 
         Caption         =   "开机启动"
      End
      Begin VB.Menu exit 
         Caption         =   "退出软件"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Con As Integer '连接数
Dim ys As Integer  '延时
Dim SVR_state As Boolean '服务器状态
'下面为托盘图标需要
Dim WithEvents m As NotifyBar
Attribute m.VB_VarHelpID = -1


Private Sub Command1_Click() '全选/全不选
Dim j%, k%
If List1.Selected(0) = False Then
    For j = 0 To List1.ListCount - 1
        
        List1.Selected(j) = True
    Next j
Else
    For k = 0 To List1.ListCount - 1
        
        List1.Selected(k) = False
    
    Next k
End If
End Sub

Private Sub Command2_Click() '踢人
Dim g%
For g = 1 To Con
DoEvents
If List1.Selected(g) = True And Connect1(g).State = 7 Then
    Connect1(g).SendData JiaMi("Scut" & "系统提示：您已经被管理员请出房间！" & " (" & Now & ") ", "")
    DoEvents
    Text2 = Text2 & "系统提示：" & List1.List(g) & "已经被管理员请出房间！" & " (" & Now & ") " & vbCrLf & vbCrLf
    Connect1(g).Close
    List1.List(g) = List1.List(g) & " [离线]"
    ys = 3
End If
Next g
End Sub

Private Sub exit_Click()
Dim yn As Integer
yn = MsgBox("确定要退出聊天室服务器吗？退出后所有聊天将断开！", vbOKCancel + vbInformation, "提示")
If yn = vbOK Then End
End Sub

Private Sub m_NotifyClick(NClickClass As NotifyClickClass, ByVal x As Long, ByVal y As Long) '托盘图标鼠标事件
Select Case NClickClass
    Case NCL_LeftButtonClick
        Me.WindowState = 0
        Me.Show
    Case NCL_RightButtonClick
        Me.PopupMenu Control '弹出菜单
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
WindowState = vbMinimized
Me.Visible = False
m.NotifyMsgBox "聊天室服务器已经最小化到托盘！", "提示", 1
End Sub

Private Sub List1_DblClick()
If Connect1(List1.ListIndex).State = 7 Then
    Dim oneQ As New ONE
    oneQ.Show
    oneQ.Caption = "与" & List1.Text & "的对话！"
End If
End Sub

Private Sub S_run_Click() '开机启动
Dim wshshell As New IWshRuntimeLibrary.wshshell
Dim mypcname As String
wshshell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Jarry_LEO", App.Path & "\" & App.EXEName & ".exe"
MsgBox "设置开机启动成功！若文件被移动后则失效，须重新设置！", vbOKOnly + 64, "提示"
End Sub

Private Sub Start_Server_Click() '开始服务
listen1.Listen
SVR_state = True
End Sub

Private Sub Stop_Server_Click() '停止服务
Dim g%
SVR_state = False
listen1.Close
For g = 1 To Con
    DoEvents
    If Connect1(g).State = 7 Then
        Connect1(g).SendData JiaMi("Scut" & "系统提示：您已与服务器断开连接！" & " (" & Now & ") " & vbCrLf, "")
        DoEvents
        Connect1(g).Close
        List1.List(g) = List1.List(g) & " [离线]"
        ys = 3
    End If
Next g
End Sub


Private Sub Command3_Click() '发送聊天信息
Dim g%
If Text3 = "" Then GoTo rr
For g = 1 To Con
    DoEvents
    If Connect1(g).State = 7 Then
        Connect1(g).SendData JiaMi("Spek" & "admin" & " (" & Now & ") " & " 说：" & vbCrLf & Text3, "")
    End If
Next g

Text2 = Text2 & "admin" & " (" & Now & ") " & "说：" & vbCrLf & Text3 & vbCrLf & vbCrLf
Text3 = ""
rr:
End Sub


Private Sub Connect1_DataArrival(index As Integer, ByVal bytesTotal As Long) '接收信息
Dim tempS As String
Dim ii As Integer
Dim cc As String
Dim g%
Dim frm As Form
    Connect1(index).GetData tempS  '读取数据
    tempS = JieMi(tempS, "")
    
Select Case Left(tempS, 4) '根据信息头4个字节分类

Case "Name"  '进入房间的信息
    List1.List(index) = Right(tempS, Len(tempS) - 4)
    Text2 = Text2 & "系统提示：" & Right(tempS, Len(tempS) - 4) & "进入房间！" & " (" & Now & ") " & vbCrLf & vbCrLf
    m.NotifyMsgBox Right(tempS, Len(tempS) - 4) & "进入房间！", "系统提示", 1
    Exit Sub
    
Case "Spek" '说话信息
    Text2 = Text2 & List1.List(index) & " (" & Now & ") " & " 说：" & vbCrLf & Right(tempS, Len(tempS) - 4) & vbCrLf & vbCrLf
    For g = 1 To Con  '转发给所有人
        DoEvents
        If Connect1(g).State = 7 Then
            Connect1(g).SendData JiaMi("Spek" & List1.List(index) & " (" & Now & ") " & " 说：" & vbCrLf & Right(tempS, Len(tempS) - 4), "")
        End If
    Next g
End Select

If Left(tempS, 2) = "To" Then  '私聊信息
    ii = Val(Mid(tempS, 3, 2))
    cc = "与" & List1.List(index) & "的对话！"
    If ii = 0 Then
        For Each frm In Forms
            If frm.Caption = cc Then
                frm.Text1 = frm.Text1 & List1.List(index) & " (" & Now & ") " & "说：" & vbCrLf & Right(tempS, Len(tempS) - 4) & vbCrLf & vbCrLf
                Exit Sub
            End If
        Next
        
        Dim oneQ As New ONE
        oneQ.Show
        oneQ.Caption = "与" & List1.List(index) & "的对话！"
        oneQ.Text1 = List1.List(index) & " (" & Now & ") " & "说：" & vbCrLf & Right(tempS, Len(tempS) - 4) & vbCrLf & vbCrLf
    
    Else
        Connect1(ii).SendData JiaMi("To" & Format(index, "00") & Right(tempS, Len(tempS) - 4), "")
    End If
End If
End Sub
Private Sub Form_Load()
If App.PrevInstance = True Then MsgBox "程序已经在运行!", vbCritical + vbOKOnly, "提示": End

'这段的作用是在托盘新建一个图标
Set m = New NotifyBar
m.Icon = Me.Icon.Handle
m.ToolTipText = "聊天室服务器！"
m.NotifyBoxVisible = True
'初始化端口
Con = 0
SVR_state = True
    listen1.Close
    listen1.LocalPort = 25627
    listen1.Listen
End Sub

Private Sub listen1_ConnectionRequest(ByVal requestID As Long) '收到连接请求
Dim z%
For z = 1 To Con    '查找空闲连接
    If Connect1(z).State <> 7 Then
        Connect1(z).Accept requestID
        GoTo ff
    End If
Next z
'一对多连接要点：一个winsock只侦听。另一个winsock数组接收连接请求，分配端口号递增！

Con = Con + 1    '没有空闲连接，创建新连接
Load Connect1(Con)
With Connect1(Con)
    .LocalPort = 25627 + Con
    .Close
End With
Connect1(Con).Accept requestID
ff:
ys = 3

End Sub

Private Sub Text2_Change()
Text2.SelStart = Len(Text2)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command3_Click
    KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer() '连接管理
Dim listA As String
Dim onL As Integer
Dim i%, x%, y%, t%
Label2 = "连 接 数：" & Con

If SVR_state = True Then
    Label1 = "运行状态：正常"
    Start_Server.Enabled = False
    Stop_Server.Enabled = True
Else
    Label1 = "运行状态：停止"
    Start_Server.Enabled = True
    Stop_Server.Enabled = False
End If

For i = 1 To Con
    If Connect1(i).State > 7 Then
        Connect1(i).Close
        Text2 = Text2 & "系统提示：" & List1.List(i) & "离开房间！" & " (" & Now & ") " & vbCrLf & vbCrLf
        m.NotifyMsgBox List1.List(i) & "离开房间！", "系统提示", 1
        List1.List(i) = List1.List(i) & " [离线]"
        ys = 3
    End If
    If Connect1(i).State = 7 Then onL = onL + 1
Next i
Label3 = "在 线 数：" & onL
If ys > 0 Then ys = ys - 1

If ys = 1 Then
    For x = 0 To List1.ListCount - 1
        listA = listA & "List" & List1.List(x)
    Next x
    
    For y = 1 To Con
        If Connect1(y).State = 7 Then
            DoEvents
            Connect1(y).SendData JiaMi(listA, "")
        End If
    Next y
End If
End Sub
