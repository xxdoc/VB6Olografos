Attribute VB_Name = "Olografoscode"
Option Explicit

Const Masculin = 0
Const Feminin = 1
Const Neutral = 2

Dim N1(1 To 3, 1 To 9) As String
Dim N2(1 To 4) As String
Dim N3(1 To 4) As String

Public Function OloTriada(s As String, Gender As Integer) As String
  Dim b As Integer
  Dim i As Integer
  Dim i2 As Integer
  Dim Result As String
  Dim h As String
  
  Result = ""
  For b = 1 To 3
    i = Asc(Mid(s, b, 1)) - 48
    '������ �� 48 ����� asc(0)=48
    If i <> 0 Then
      h = N1(4 - b, i)
      
      Select Case b
        Case 1
          If i <> 1 Then
            If Gender = Masculin Then
              h = h + "��"
            ElseIf Gender = Feminin Then
              h = h + "��"
            Else
              h = h + "�"
            End If
          End If
        Case 2
          If i = 1 Then
            i2 = Asc(Mid(s, 3, 1)) - 48
            If i2 = 1 Or i2 = 2 Then
              If i2 = 1 Then
                h = "������"
              ElseIf i2 = 2 Then
                h = "������"
              End If
              Result = Result + h
              Exit For
            End If
          End If
        Case 3
          If i = 1 Then
            If Gender = Feminin Then
              h = "���"
            End If
          ElseIf i = 3 Then
            If Gender = Neutral Then
              h = h + "��"
            Else
              h = h + "���"
            End If
          ElseIf i = 4 Then
            If Gender = Neutral Then
              h = h + "�"
            Else
              h = h + "��"
            End If
          End If
      End Select
      
      Result = Result + h
      If b <> 3 Then Result = Result + " "
    End If
  Next b
  OloTriada = Result
End Function

Public Function Olografos(Num As Double, Gender As Integer) As String
  Dim iNum As Double
  Dim s As String
  Dim Result As String
  Dim h As String
  Dim b As Integer
  Dim Gen As Integer
  Dim triada As Integer
  Dim Decimicals As Boolean
    
  InitNames
  iNum = Int(Num)
  Decimicals = iNum <> Num
  If iNum = 0 Then
    Result = "�����"
    GoTo GiveResult
  End If
  s = Trim(Str(iNum))
  If Len(s) Mod 3 <> 0 Then s = String(3 - (Len(s) Mod 3), "0") + s
  b = 0
  Result = ""
  While b < Len(s)
    triada = (Len(s) - b) \ 3
    ' � ������ ��������� ��� �����
    ' �.� � ���� 2 ����� � ������ ��� ��������
    If triada = 1 Then
      Gen = Gender
    ElseIf triada = 2 Then
      Gen = Feminin
    Else
      Gen = Neutral
    End If
    If triada = 2 And Val(Mid(s, b + 1, 3)) = 1 Then
      Result = Result + "������ "
    Else
      Result = Result + OloTriada(Mid(s, b + 1, 3), Gen) + " "
      If triada <> 1 Then
        h = N2(triada - 1)
        If triada > 2 Then
          If Val(Mid(s, b + 1, 3)) = 1 Then
            h = h + "�"
          Else
            h = h + "�"
          End If
        End If
        Result = Result + h + " "
      End If
    End If
    
    b = b + 3
  Wend
GiveResult:
  Mid(Result, 1, 1) = UCase(Mid(Result, 1, 1))
  Olografos = RTrim(Result)
End Function

Public Function Olografos���(Num As Double) As String
  Olografos��� = Olografos(Num, Feminin) + " ���"
End Function
Public Sub InitNames()
  N1(1, 1) = "���"
  N1(1, 2) = "���"
  N1(1, 3) = "��"
  N1(1, 4) = "������"
  N1(1, 5) = "�����"
  N1(1, 6) = "���"
  N1(1, 7) = "����"
  N1(1, 8) = "����"
  N1(1, 9) = "�����"
  N1(2, 1) = "����"
  N1(2, 2) = "������"
  N1(2, 3) = "�������"
  N1(2, 4) = "�������"
  N1(2, 5) = "�������"
  N1(2, 6) = "������"
  N1(2, 7) = "���������"
  N1(2, 8) = "�������"
  N1(2, 9) = "��������"
  N1(3, 1) = "�����"
  N1(3, 2) = "�������"
  N1(3, 3) = "��������"
  N1(3, 4) = "���������"
  N1(3, 5) = "���������"
  N1(3, 6) = "�������"
  N1(3, 7) = "��������"
  N1(3, 8) = "��������"
  N1(3, 9) = "���������"
  
  N2(1) = "��������"
  N2(2) = "����������"
  N2(3) = "�������������"
  N2(4) = "��������������"
  
  N3(1) = "������"
  N3(2) = "��������"
  N3(3) = "��������"
  N3(4) = "������� ��������"
End Sub
