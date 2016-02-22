Attribute VB_Name = "���ܽ���ģ��"
'����ԭ��
'�������10�����ظ���������Կ
'��������ת��Ϊ���֣���ת��Ϊ������Ķ����ƣ�
'��ԭ�ĵ��ַ�ȫ��ת��Ϊ����
'��ԭ�İ�������Ķ����ƺ���Կ�����ֽ���ѭ���滻��1Ϊ�滻��0Ϊ���滻
'����Կ��һ�������滻Ϊ������ֲ���ת��Կ���˶�Ŀ��Ȼ��������Ŀ�ͷ


Private Function ShuZi(ByVal x As String) As String '�ַ�ת��Ϊ����
If Asc(x) < 0 Then
    x = Hex(Asc(x))                         '����ת��
    ShuZi = Format(CDec("&H" & Left(x, 2)) - 160, "00") & Format(CDec("&H" & Right(x, 2)) - 160, "00") '��ȡ������λ��
Else
 
    ShuZi = Format(Asc(x) + 9527, "0000")   'Ӣ�����ַ���ת��Ϊ��λ���֡�
End If
End Function
Private Function WenZi(ByVal x As String) As String '����ת��Ϊ�ַ�
On Error Resume Next
If IsNumeric(x) = False Then Exit Function
If x < 9000 Then
    WenZi = Chr((CLng(Left(x, 2)) + 160) * &H100 + CLng(Right(x, 2)) + 160) '��λ��ת��Ϊ����
Else
    WenZi = Chr(x - 9527)
End If
End Function

Private Function Zzhu(ByVal x As String) As String '�����ַ�ת��Ϊ����
For i = 1 To Len(x)
    Zzhu = Zzhu + ShuZi(Mid(x, i, 1))
Next
End Function
Private Function Szhu(ByVal x As String) As String '��������ת��Ϊ�ַ�
For i = 1 To Len(x) Step 4
    Szhu = Szhu + WenZi(Mid(x, i, 4))
Next
End Function
Private Function MiWen() As String                 '���10λ��Կ��0-9�����ظ�
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
Private Function Bin2(ByVal x As String, p As String) As String  '������ת��Ϊ������
Dim s As String, l As String
Dim k%, j%, m%
For k = 1 To Len(x)
    For j = 1 To Len(x)
        m = Val(Mid(p, j Mod 10 + 1, 1))
        l = Val(Mid(x, k, 1)) Xor m Xor Val(Mid(x, j, 1)) '����ת��������ѭ�����������Կ�����������
    Next j
    s = s + Bin1(l) '�ѽ�����ж�����ת��
Next k
Bin2 = s
End Function
Private Function Bin1(ByVal x As String) As String  '��������ת��Ϊ������
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
Public Function JiaMi(ByVal x As String, ByVal y As String) As String '���ܹ��̣�X ΪҪ���ܵ����֡�YΪ���룬���Բ��裩
Dim PW As String
Dim BPW As String
Dim st As String
Dim tem As Integer
Dim ZZ As String
PW = MiWen
If Len(y) = 0 Then  '���û���������룬��ô������������Ϊ�����Կ�Ķ����ƣ�����Ϊ����Ķ�����
    BPW = Bin2(Zzhu(PW), PW)
Else
    BPW = Bin2(Zzhu(y), PW)
End If
st = Zzhu(x)        'ԭ��ת��Ϊ����
For i = 1 To Len(st)
    tem = tem + 1
    If tem > Len(BPW) Then tem = 1
    If Mid(BPW, tem, 1) = 1 Then    '�����������������0��1��ԭ�ĵ������滻��ԭ����������Կ�е����λ�õ�����
        ZZ = ZZ & Mid(PW, Mid(st, i, 1) + 1, 1)
    Else
        ZZ = ZZ & (Val(Mid(st, i, 1)) Xor 1) '�����һ�¡������ƽ��Ѷ�
    End If
Next i
Randomize
JiaMi = Int(Rnd * 10) & Right(StrReverse(PW), 9) & ZZ '��һ��������ִ��涪ʧ�ķ�ת��Կ���ַ����ֶ�ͷ�����۷�
End Function
Public Function JieMi(ByVal x As String, ByVal y As String) As String '���ܹ���
On Error Resume Next
Dim PW As String
Dim BPW As String
Dim st As String
Dim tem As Integer
Dim ZZ As String
PW = StrReverse(Mid(x, 2, 9)) & FindLost(Mid(x, 2, 9))  '��ȡ��Կ
If Len(y) = 0 Then
    BPW = Bin2(Zzhu(PW), PW)   '����ת���ɼ��ܶ�����
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
Private Function FindLost(ByVal x As String) As String '���Ҷ�ʧ������
For u = 0 To 9
    If InStr(1, x, u) = 0 Then FindLost = u
Next
End Function
