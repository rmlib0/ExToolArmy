Attribute VB_Name = "function_reparsing"
' --------------------------------------------------------------------------------
'
'       Title:      ExToolArmy.xlam
'
'       Purpose:    ����� �������, ��� ����������� ����������� � �����
'                   Microsoft Office 2017 Excel
'
'       License:    BearLic v.0.1a
'
'       DateTime:   05/14/2020 14:00 GMT+3
'
'       Author:     rmlib.null (OVSYANNIKOV Mikhail Mikhailovich)
'                              (���������� ������ ����������)
'
'       Contacts:   rmlib.null@gmail.com
'                   +7 (918) 518 40-62
'
' --------------------------------------------------------------------------------

'
' StrPrepare
'
' ������� ���������� ������
'
Function StrPrepare(m As String) As String

    ' ������ ������� ����� � ������
    m = Trim(m)
    
    ' ������ ������� �������
    Do While InStr(1, m, Space(2), 1) <> 0
        m = Replace(m, Space(2), Space(1), vbTextCompare)
    Loop
    
    StrPrepare = m
    
End Function

'
' DTW_Month
'
' ������� �������� ������ � �����
'
Public Function DTW_Month(m As Integer, pcase As Integer) As String

    Dim arrJoin(1 To 12) As Variant
    
    arrJoin(1) = Array("������", "������", "������", "������", "�������", "������")
    arrJoin(2) = Array("�������", "�������", "�������", "�������", "��������", "�������")
    arrJoin(3) = Array("����", "�����", "�����", "����", "������", "�����")
    arrJoin(4) = Array("������", "������", "������", "������", "�������", "������")
    arrJoin(5) = Array("���", "���", "���", "���", "����", "���")
    arrJoin(6) = Array("����", "����", "����", "����", "�����", "����")
    arrJoin(7) = Array("����", "����", "����", "����", "�����", "����")
    arrJoin(8) = Array("������", "�������", "�������", "������", "��������", "�������")
    arrJoin(9) = Array("��������", "��������", "��������", "��������", "���������", "��������")
    arrJoin(10) = Array("�������", "�������", "�������", "�������", "��������", "�������")
    arrJoin(11) = Array("������", "������", "������", "������", "�������", "������")
    arrJoin(12) = Array("�������", "�������", "�������", "�������", "��������", "�������")

    DTW_Month = arrJoin(m)(pcase - 1)
    
End Function

'
' FaceBit
'
' ������� ����������� �������� � ������
'
Function FaceBit(m As String) As Boolean
    
    arrAbsence = Array("������������", _
                    "������", _
                    "���������", _
                    "�������", _
                    "�������", _
                    "�������", _
                    "�����", _
                    "���", _
                    "�/�", _
                    "�����", _
                    "�����������", _
                    "������", _
                    "�������", _
                    "�� �������", _
                    "�����", _
                    "�����", _
                    "��������� ����", _
                    "�������", _
                    "2 ���", _
                    "��������", _
                    "������" _
                    )
                    
    m = StrPrepare(m)
    
    For i = 0 To UBound(arrAbsence)
        If arrAbsence(i) = m Then
            FaceBit = True
            Exit For
        End If
    Next i

End Function

'
' RprtPrep
'
' ������� ���������� ������ ��� ����������� �������
'
Function RprtPrep(m As Variant, t As Range) As Variant
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim arrJoin() As Variant
    Dim arrPeople() As Variant
    Dim arrBuff() As Variant

    Dim NumRows As Integer
    Dim NumColumns As Integer
    
    m = StrConv(CStr(Trim(m)), vbLowerCase)
    
    ' �������� ���������� �����
    NumRows = t.Rows.Count
    
    ' �������� ���������� ��������
    NumColumns = t.Columns.Count
    
    ' ���������� ������� ��������
    ReDim arrBuff(0 To NumColumns - 1)
    ReDim arrJoin(0 To NumRows - 1)

    
    For i = 1 To NumRows
        If StrConv(CStr(Trim(t(i, 1).Value)), vbLowerCase) = m Then
            arrBuff(0) = t(i, 2).Value
            arrBuff(1) = t(i, 3).Value
            arrBuff(2) = t(i, 4).Value
            
            ' �������� ���� �����������
            ' ���� �� ������ ��� �� ����� ����, ����� ����������
            If arrBuff(2) <> vbEmpty Or Format(arrBuff(2), "yyyy") > 2000 Then
                arrJoin(j) = ������_��������_������(CStr(arrBuff(0))) + " " + ��������(CStr(arrBuff(1)), 1) + " �� " + CStr(arrBuff(2))
                
                ' ���� ���� ����������� ������ ��� ����� ����
                Else:
                    arrJoin(j) = ������_��������_������(CStr(arrBuff(0))) + " " + ��������(CStr(arrBuff(1)), 1)
            End If
            
            j = j + 1
            
            Else:
                RprtPrep = ""
        End If
    Next i
    
    ' ���������� ������� �������
    ReDim arrPeople(0 To j - 1)
    
    ' ��������� ������
    For k = 0 To UBound(arrPeople)
        arrPeople(k) = CStr(arrPeople(k)) + CStr(arrJoin(k))
    Next k

    ' ������� �������� �������, �������� ��������������
    RprtPrep = StrConv(CStr(m), vbLowerCase) + " " + _
                Chr(151) + " " + CStr(j) + " ���. (" + _
                Join(arrPeople, ", ") + ")"

End Function

'
' RprtUnion
'
' ������� ������ ��� ����������� �������
'
Function RprtUnion(d As Variant, t As Range) As Variant

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    Dim arrBuff() As Variant

    Dim NumRows As Integer

    d = StrConv(CStr(Trim(d)), vbLowerCase)
    
    ' �������� ���������� �����
    NumRows = t.Rows.Count
    
    ' ���������� ������� �������
    ReDim arrBuff(1 To NumRows)
    
    j = 1
    
    For i = 1 To NumRows
        If StrConv(CStr(Trim(t(i, 1).Value)), vbLowerCase) <> d Then
        
            ' ���� ������ �� �����������, �� ������� �������� � ������ � ��������� �������
            arrBuff(j) = t(i, 1).Value
            j = j + 1
            
        End If
    Next i
    
    ReDim Preserve arrBuff(1 To j)
    
    RprtUnion = Join(arrBuff, Chr(10))
    
End Function

Private Function fRole(m As String) As String
    
    m = StrPrepare(m)
    
    arrRole_1 = Array("�������� ���������", _
                        "����������� ��������� ��������� �� ���", _
                        "����������� ��������� ��������� �� ����������", _
                        "�������� ��������� ��������� �� ������ ����������", _
                        "��������� ����� - ����������� ��������� ���������", _
                        "����������� ���������� �����", _
                        "�������� ����", _
                        "����������� ��������� ���� �� ���", _
                        "�������� ������", _
                        "�������� �������", _
                        "�������� ������ - ������� ������ �� �������", _
                        "��������� ����� - �������� ������" _
                        )
    For i = 0 To UBound(arrRole_1)
        If m = arrRole_1(i) Then
            fRole = sRole_1
            Exit For
        End If
    Next i
    
End Function

Private Function ����_�_�������(��������� As String) As String
    Dim ���� As Integer
    
    ��������� = StrPrepare(���������)
    
        If ��������� = "�������� ���������" Or _
            ��������� = "����������� ��������� ��������� �� ���" Or _
            ��������� = "����������� ��������� ��������� �� ����������" Or _
            ��������� = "�������� ��������� ��������� �� ������ ����������" Or _
            ��������� = "��������� ����� - ����������� ��������� ���������" Or _
            ��������� = "����������� ���������� �����" Or _
            ��������� = "�������� ����" Or _
            ��������� = "����������� ��������� ���� �� ���" Or _
            ��������� = "�������� ������" Or _
            ��������� = "�������� �������" Or _
            ��������� = "�������� ������ - ������� ������ �� �������" Or _
            ��������� = "��������� ����� - �������� ������" Then
                ���� = 1
            ElseIf ��������� = "����������� ��������� ������ - �������� ���������" Or _
                    ��������� = "��������" Or _
                    ��������� = "������� ������" Or _
                    ��������� = "����������� ��������� ������ - �������� ��������" Or _
                    ��������� = "�������� ���������" Or _
                    ��������� = "�������� ��������" Or _
                    ��������� = "�������� ��������� (������� �����������)" Then
                ���� = 2
            Else:
                ���� = 3
        End If
    Select Case ����
        Case 1
            ����_�_������� = "��������� �������"
        Case 2
            ����_�_������� = "�������� ���������� �������, ����������"
        Case 3
            ����_�_������� = "����������"
    End Select
End Function

Function ���_�������_��_������(���� As String, ����� As String, ��������� As String, ��������� As String, ������ As String, ��� As String) As String
    
    ��������� = Trim(���������)
    Do While InStr(1, ���������, Space(2), 1) <> 0
        ��������� = Replace(���������, Space(2), Space(1), vbTextCompare)
    Loop
    
    ���_�������_��_������ = ( _
                                ��������� + Chr(32) + _
                                ������_��������_������(������) + Chr(32) + _
                                ��������(StrConv(����������(���), vbProperCase), 1) + Chr(32) + Chr(150) + Chr(32) + _
                                ����_�_�������(���������) + ";")
End Function

