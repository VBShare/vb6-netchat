VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��½������"
   ClientHeight    =   3600
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4215
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2126.998
   ScaleMode       =   0  'User
   ScaleWidth      =   3957.657
   StartUpPosition =   2  '��Ļ����
   Begin �����ҿͻ���.jcbutton cmdOK 
      Height          =   435
      Left            =   2640
      TabIndex        =   10
      Top             =   1020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "��¼"
   End
   Begin �����ҿͻ���.jcbutton cmdCancel 
      Height          =   435
      Left            =   300
      TabIndex        =   9
      Top             =   1020
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����"
   End
   Begin VB.Frame Frame1 
      Caption         =   "����������"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1440
         TabIndex        =   6
         Text            =   "25627"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1440
         TabIndex        =   4
         Text            =   "jarryleo.vicp.cc"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�������˿ڣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "������������"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1530
      MaxLength       =   12
      TabIndex        =   1
      Top             =   495
      Width           =   2325
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������һ�㲻��Ҫ���ģ�"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1800
      TabIndex        =   8
      Top             =   3240
      Width           =   2160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ȡһ���������֪������ݵ����ƣ�"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   3060
   End
   Begin VB.Line Line1 
      X1              =   112.674
      X2              =   3830.899
      Y1              =   992.599
      Y2              =   992.599
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "�û�����(&U):"
      Height          =   270
      Index           =   0
      Left            =   345
      TabIndex        =   0
      Top             =   510
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���˵�½��ť�Ĵ��롣������ڵĴ��붼�ǽ���������ȡIP��
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADESCRIPTION_LEN = 256
Private Const WSASYS_STATUS_LEN = 128

Private Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLength As Integer
   hAddrList As Long
End Type

Private Type WSAData
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADESCRIPTION_LEN) As Byte
   szSystemStatus(0 To WSASYS_STATUS_LEN) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type
Private Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Any, ByVal _
byteslen As Integer, addrtype As Integer) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal _
        wVersionRequired&, lpWSADATA As WSAData) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal _
        hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, _
        ByVal hpvSource&, ByVal cbCopy&)
Dim Nam As String

Function hibyte(ByVal wParam As Integer)   ' ע�ͣ���������ĸ�λ
   hibyte = wParam \ &H100 And &HFF&
End Function

Function lobyte(ByVal wParam As Integer)   ' ע�ͣ���������ĵ�λ
   lobyte = wParam And &HFF&
End Function

Function SocketsInitialize()
   Dim WSAD As WSAData
   Dim iReturn As Integer
   Dim sLowByte As String, sHighByte As String, sMsg As String
   
   iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
   
   If iReturn <> 0 Then
      MsgBox "Winsock.dll û�з�Ӧ."
      End
   End If
   
   If lobyte(WSAD.wVersion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wVersion) = WS_VERSION_MAJOR And hibyte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      sHighByte = Trim$(Str$(hibyte(WSAD.wVersion)))
      sLowByte = Trim$(Str$(lobyte(WSAD.wVersion)))
      sMsg = "Windows Sockets�汾 " & sLowByte & "." & sHighByte
      sMsg = sMsg & " ����winsock.dll֧�� "
      MsgBox sMsg
      End
   End If
   
   If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
      sMsg = "���ϵͳ��Ҫ������Sockets��Ϊ "
      sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD))
      MsgBox sMsg
      End
   End If
   
End Function

Sub SocketsCleanup()
   Dim lReturn As Long
   
   lReturn = WSACleanup()
   
   If lReturn <> 0 Then
      MsgBox "Socket���� " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
      End
   End If
End Sub

Private Sub cmdCancel_Click() '��ʾ/��������
Me.Height = 6100 - Me.Height
End Sub

Private Sub cmdOK_Click() '��½
If IsNumeric(Text2) = False Or Val(Text2) < 0 Or Val(Text2) > 65535 Then
    MsgBox "�˿ڸ�ʽ����", vbOKOnly + 16, "���棡"
    Exit Sub
End If
If getip(Text1) = "" Then
    MsgBox "������IP��ַ�������󣡿����Ƿ����������ߣ����Ժ����ԣ�", vbOKOnly + 64, "��ʾ��"
    Exit Sub
End If

SVR_IP = getip(Text1)
SVR_PORT = Val(Text2)
USER_NAME = txtUserName
If USER_NAME = "" Then USER_NAME = "��������"
If LCase(USER_NAME) = "admin" Then USER_NAME = "ð��admin"
FormMain.Show
Unload Me
End Sub

Sub Form_Load()
    Me.Height = 2100
'ע��:     ��ʼ��Socket
    SocketsInitialize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'ע��:     ���Socket
    SocketsCleanup
End Sub
Private Function getip(name As String) As String '����������ȡ������IP��ַ
   Dim hostent_addr As Long
   Dim Host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   
   hostent_addr = gethostbyname(name)
   
   If hostent_addr = 0 Then
      getip = ""                     'ע�ͣ����������ܱ�����
      Exit Function
   End If
   
   RtlMoveMemory Host, hostent_addr, LenB(Host)
   RtlMoveMemory hostip_addr, Host.hAddrList, 4
   
   ReDim temp_ip_address(1 To Host.hLength)
   RtlMoveMemory temp_ip_address(1), hostip_addr, Host.hLength
   
   For i = 1 To Host.hLength
      ip_address = ip_address & temp_ip_address(i) & "."
   Next
   ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
   
   getip = ip_address

End Function

