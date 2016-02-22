Attribute VB_Name = "加密解密模块"
'加密原理：
'随机生成10个不重复的数字密钥
'把密码先转化为数字，再转化为处理过的二进制；
'把原文的字符全部转化为数字
'把原文按照密码的二进制和密钥的数字进行循环替换。1为替换，0为不替换
'把密钥第一个数字替换为随机数字并反转密钥掩人耳目，然后放在密文开头


Private Function ShuZi(ByVal x As String) As String '字符转化为数字
If Asc(x) < 0 Then
    x = Hex(Asc(x))                         '中文转化
    ShuZi = Format(CDec("&H" & Left(x, 2)) - 160, "00") & Format(CDec("&H" & Right(x, 2)) - 160, "00") '求取汉字区位码
Else
 
    ShuZi = Format(Asc(x) + 9527, "0000")   '英文数字符号转化为四位数字。
End If
End Function
Private Function WenZi(ByVal x As String) As String '数字转化为字符
On Error Resume Next
If IsNumeric(x) = False Then Exit Function
If x < 9000 Then
    WenZi = Chr((CLng(Left(x, 2)) + 160) * &H100 + CLng(Right(x, 2)) + 160) '区位码转换为汉字
Else
    WenZi = Chr(x - 9527)
End If
End Function

Private Function Zzhu(ByVal x As String) As String '所有字符转化为数字
For i = 1 To Len(x)
    Zzhu = Zzhu + ShuZi(Mid(x, i, 1))
Next
End Function
Private Function Szhu(ByVal x As String) As String '所有数字转化为字符
For i = 1 To Len(x) Step 4
    Szhu = Szhu + WenZi(Mid(x, i, 4))
Next
End Function
Private Function MiWen() As String                 '随机10位密钥（0-9）不重复
Dim b%, c%
Dim a As String
Dim tmp As String
For b = 1 To 10
tt:
    DoEvents
    Randomize
    tmp = Chr(Int(Rnd(1) * 10 + 48))
    For c = 1 To Len(a)
        If Mid(a, c, 1) = tmp Then GoTo tt
    Next c
    a = a + tmp
Next b
    MiWen = a
End Function
Private Function Bin2(ByVal x As String, p As String) As String  '把密码转化为二进制
Dim s As String, l As String
Dim k%, j%, m%
For k = 1 To Len(x)
    For j = 1 To Len(x)
        m = Val(Mid(p, j Mod 10 + 1, 1))
        l = Val(Mid(x, k, 1)) Xor m Xor Val(Mid(x, j, 1)) '密码转化的数字循环与自身和密钥进行异或运算
    Next j
    s = s + Bin1(l) '把结果进行二进制转换
Next k
Bin2 = s
End Function
Private Function Bin1(ByVal x As String) As String  '单个数字转化为二进制
Dim i As Integer
Dim a As Integer
Dim b As String
    a = x
    Do
        i = a Mod 2
        a = a \ 2
        b = i & b
    Loop While a > 0
    Bin1 = b
End Function
Public Function JiaMi(ByVal x As String, ByVal y As String) As String '加密过程（X 为要加密的文字。Y为密码，可以不设）
Dim PW As String
Dim BPW As String
Dim st As String
Dim tem As Integer
Dim ZZ As String
PW = MiWen
If Len(y) = 0 Then  '如果没有设置密码，那么二进制密密码为随机密钥的二进制，否则为密码的二进制
    BPW = Bin2(Zzhu(PW), PW)
Else
    BPW = Bin2(Zzhu(y), PW)
End If
st = Zzhu(x)        '原文转化为数字
For i = 1 To Len(st)
    tem = tem + 1
    If tem > Len(BPW) Then tem = 1
    If Mid(BPW, tem, 1) = 1 Then    '按二进制密码里面的0和1把原文的数字替换成原文数字在密钥中的这个位置的数字
        ZZ = ZZ & Mid(PW, Mid(st, i, 1) + 1, 1)
    Else
        ZZ = ZZ & (Val(Mid(st, i, 1)) Xor 1) '再异或一下。增加破解难度
    End If
Next i
Randomize
JiaMi = Int(Rnd * 10) & Right(StrReverse(PW), 9) & ZZ '用一个随机数字代替丢失的反转密钥数字放入字段头，障眼法
End Function
Public Function JieMi(ByVal x As String, ByVal y As String) As String '解密过程
On Error Resume Next
Dim PW As String
Dim BPW As String
Dim st As String
Dim tem As Integer
Dim ZZ As String
PW = StrReverse(Mid(x, 2, 9)) & FindLost(Mid(x, 2, 9))  '获取密钥
If Len(y) = 0 Then
    BPW = Bin2(Zzhu(PW), PW)   '密码转换成加密二进制
Else
    BPW = Bin2(Zzhu(y), PW)
End If
st = Right(x, Len(x) - 10)
For i = 1 To Len(st)
    tem = tem + 1
    If tem > Len(BPW) Then tem = 1
    If Mid(BPW, tem, 1) = 1 Then
        For k = 1 To 10
            If Mid(PW, k, 1) = Mid(st, i, 1) Then ZZ = ZZ & CStr(k - 1)
        Next k
    Else
        ZZ = ZZ & (Val(Mid(st, i, 1)) Xor 1)
    End If
Next i
JieMi = Szhu(ZZ)
Debug.Print PW
End Function
Private Function FindLost(ByVal x As String) As String '查找丢失的数字
For u = 0 To 9
    If InStr(1, x, u) = 0 Then FindLost = u
Next
End Function
