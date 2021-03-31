Attribute VB_Name = "functions"
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

Public Function ����������(������ As Range) As Double
    ���������� = ������.Interior.Color
End Function

Public Function ���������(�������� As String, ����� As Integer) As Variant
    Dim buffLastName As String
    Dim buffFirstName As String
    Dim buffMidName As String
    Dim buffPostMidName As String
    Dim buffSplit() As String
        
    Dim gradLastName As String
    Dim gradFirstName As String
    Dim gradMidName As String
    
    ' ������ ������� ������� �� ������� ������
    Do While InStr(1, ��������, Space(2), 1) <> 0
        �������� = Replace(��������, Space(2), Space(1), vbTextCompare)
    Loop

    ' ������� ������ � ������������ " "
    buffSplit() = Split(Trim(��������), Chr(32))
    
    ' �������� �������� - �������
    buffLastName = buffSplit(0)
    ' �������� �������� - ���
    buffFirstName = buffSplit(1)
    ' �������� �������� - ��������
    buffMidName = buffSplit(2)
    ' TODO
    ' �������� �������� - ���� �������� ���� ����
    ' buffPostMidName = buffSplit(3)
        
    gradLastName = ���������_�������(buffLastName, �����)
    gradFirstName = ���������_���(buffFirstName, �����)
    gradMidName = ���������_��������(buffMidName, �����)
    
    ��������� = gradLastName & Chr(32) & gradFirstName & Chr(32) & gradMidName & Chr(32) & buffPostMidName

End Function

Public Function ���������������(�������� As String, ����� As Integer) As Variant
    Dim buffSplit() As String
    Dim buffSimpleGrade As String
    Dim buffComplexGrade As String
    Dim buffOtherGrade As String
    
    Dim flagSimpleGrade As Integer
    Dim flagComplexGrade As Integer
    Dim flagOtherGrade As Integer
         
    �������� = StrPrepare(��������)
    
    ' ������ ������� ����� � ������
    '�������� = Trim(��������)
    
    ' ������ ������� �������
    'Do While InStr(1, ��������, Space(2), 1) <> 0
    '    �������� = Replace(��������, Space(2), Space(1), vbTextCompare)
    'Loop
    
    ' ������� ������ � ������������ " "
    buffSplit() = Split(��������, Chr(32))
        
    j = 0
    flagSimpleGrade = 0
    flagComplexGrade = 0
    flagOtherGrade = 0
    
    Do While j <= UBound(buffSplit())
        If j = 0 Then
            flagSimpleGrade = 1
            ElseIf j = 2 Then
                flagOtherGrade = 1
            Else
                flagComplexGrade = 1
        End If
        j = j + 1
    Loop
    
    If flagSimpleGrade = 1 Then
        buffSimpleGrade = buffSplit(0)
            If flagComplexGrade = 1 Then
                buffSimpleGrade = buffSplit(0)
                buffComplexGrade = buffSplit(1)
                If flagOtherGrade = 1 Then
                    buffSimpleGrade = buffSplit(0)
                    buffComplexGrade = buffSplit(1)
                    buffOtherGrade = buffSplit(2)
                End If
            End If
    End If
        
        extEndingSimpleGrade = LCase(Mid(CStr(buffSimpleGrade), (Len(CStr(buffSimpleGrade)) - 2), Len(CStr(buffSimpleGrade))))
        Select Case �����
            ' ������������
            Case 1
                If flagSimpleGrade = 1 Then
                    buffSimpleGrade = buffSplit(0)
                    ��������������� = buffSimpleGrade
                    If flagComplexGrade = 1 Then
                        buffSimpleGrade = buffSplit(0)
                        buffComplexGrade = buffSplit(1)
                        ��������������� = buffSimpleGrade + Chr(32) + buffComplexGrade
                        If flagOtherGrade = 1 Then
                            buffSimpleGrade = buffSplit(0)
                            buffComplexGrade = buffSplit(1)
                            buffOtherGrade = buffSplit(2)
                            ��������������� = buffSimpleGrade + Chr(32) + buffComplexGrade + Chr(32) + buffOtherGrade
                        End If
                    End If
                End If
                
            ' �����������
            Case 2
                If flagSimpleGrade = 1 Then
                        ' ���������
                        ' ������������
                    If (extEndingSimpleGrade = "���") Then
                        ��������������� = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "���"
                        ' �����
                        ElseIf (extEndingSimpleGrade = "���") Then
                            ��������������� = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "���"
                        ' �������
                        ElseIf (extEndingSimpleGrade = "���") Then
                            ��������������� = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "���"
                        ' ���������
                        ' �������
                        ElseIf (extEndingSimpleGrade = "���") Then
                            ��������������� = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "���"
                        ' ���������
                        ElseIf (extEndingSimpleGrade = "���") Then
                            ��������������� = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "���"
                        ' ��������
                       ElseIf (extEndingSimpleGrade = "���") Then
                            ��������������� = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "��"
                        ' ��������
                        ElseIf (extEndingSimpleGrade = "���") Then
                            ��������������� = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "���"
                        ' �������
                        ElseIf (extEndingSimpleGrade = "���") Then
                            ��������������� = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "���"
                        Else: ��������������� = "!!!"
                    End If
                End If
        End Select
'    End If
    
    'If InStr(1, buffSplit(1), Chr(45), 1) Then
        ' ���� ��� ������� �� 2-� ����, ��
     '   flagFirstName = 1
     '   buffSimpleGrade() = Split(Trim(buffSplit(1)), Chr(45))
     '   j = 0
    '    i = 0
    '    Do While j <= UBound(buffSimpleGrade())
    '        buffSimpleGrade(j) = StrConv(buffSimpleGrade(j), vbProperCase)
   '         j = j + 1
    '    Loop
    '    j = 0
    '    Do While i <= UBound(buffSimpleGrade())
    '        buffSimpleGrade(i) = StrConv(buffSimpleGrade(i), vbProperCase)
    '        i = i + 1
    '    Loop
   '     i = 0
   '     Else
'            ��������������� = StrConv(Trim(buffSplit(1)), vbProperCase)
'    End If
    
    ' �������� �������� - ������ �������
    'buffSimpleGrade = buffSplit(0)
    ' �������� �������� - ��������� ������
    'buffComplexGrade = buffSplit(1)
    ' �������� �������� - ���., ����� � ������
    'buffOtherGrade = buffSplit(2)

End Function


Private Function �����������������(�������� As String) As String
    Dim ArrayCyrillic(1 To 33) As String
    Dim ArrayLatin(1 To 26) As String
    Dim i As Long
    
    ArrayCyrillic(1) = "�"
    ArrayCyrillic(2) = "�"
    ArrayCyrillic(3) = "�"
    ArrayCyrillic(4) = "�"
    ArrayCyrillic(5) = "�"
    ArrayCyrillic(6) = "�"
    ArrayCyrillic(7) = "�"
    ArrayCyrillic(8) = "�"
    ArrayCyrillic(9) = "�"
    ArrayCyrillic(10) = "�"
    ArrayCyrillic(11) = "�"
    ArrayCyrillic(12) = "�"
    ArrayCyrillic(13) = "�"
    ArrayCyrillic(14) = "�"
    ArrayCyrillic(15) = "�"
    ArrayCyrillic(16) = "�"
    ArrayCyrillic(17) = "�"
    ArrayCyrillic(18) = "�"
    ArrayCyrillic(19) = "�"
    ArrayCyrillic(20) = "�"
    ArrayCyrillic(21) = "�"
    ArrayCyrillic(22) = "�"
    ArrayCyrillic(23) = "�"
    ArrayCyrillic(24) = "�"
    ArrayCyrillic(25) = "�"
    ArrayCyrillic(26) = "�"
    ArrayCyrillic(27) = "�"
    ArrayCyrillic(28) = "�"
    ArrayCyrillic(29) = "�"
    ArrayCyrillic(30) = "�"
    ArrayCyrillic(31) = "�"
    ArrayCyrillic(32) = "�"
    ArrayCyrillic(33) = "�"
    
    ArrayLatin(1) = "a"
    ArrayLatin(2) = "b"
    ArrayLatin(3) = "c"
    ArrayLatin(4) = "d"
    ArrayLatin(5) = "e"
    ArrayLatin(6) = "f"
    ArrayLatin(7) = "g"
    ArrayLatin(8) = "h"
    ArrayLatin(9) = "i"
    ArrayLatin(10) = "j"
    ArrayLatin(11) = "k"
    ArrayLatin(12) = "l"
    ArrayLatin(13) = "m"
    ArrayLatin(14) = "n"
    ArrayLatin(15) = "o"
    ArrayLatin(16) = "p"
    ArrayLatin(17) = "q"
    ArrayLatin(18) = "r"
    ArrayLatin(19) = "s"
    ArrayLatin(20) = "t"
    ArrayLatin(21) = "u"
    ArrayLatin(22) = "v"
    ArrayLatin(23) = "w"
    ArrayLatin(24) = "x"
    ArrayLatin(25) = "y"
    ArrayLatin(26) = "z"

    
    If Len(��������) = 0 Then
        ����������������� = ""
    Else
        For i = 1 To Len(��������)
            CurrentChar = Mid(��������, i, 1)
            For m = 1 To 33
                If CurrentChar = ArrayCyrillic(m) Then
                    ConvertChar = ConvertChar & Trim(Str(m)) & Chr(32)
                    Exit For
                End If
            Next m
            For N = 1 To 26
                If CurrentChar = ArrayLatin(N) Then
                    ConvertChar = ConvertChar & "!" & Trim(Str(N)) & Chr(32)
                    Exit For
                End If
            Next N
        Next i
        ����������������� = Trim(ConvertChar)
    End If
End Function

Function ��������(�������� As String, ��� As Integer) As Variant
    
    Dim buffLastName() As String
    Dim buffFirstName() As String
    Dim buffMidName() As String
    
    Dim buffSplit() As String
       
    Dim flagFirstName As Integer
    Dim flagMidName As Integer
    
    Dim LastName As String
    Dim FirstName As String
    Dim MidName As String
    
    ' ������ ������� ����� � ������
    �������� = Trim(��������)
    
    ' ������ ������� �������
    Do While InStr(1, ��������, Space(2), 1) <> 0
        �������� = Replace(��������, Space(2), Space(1), vbTextCompare)
    Loop
    
    ' ������� ������ � ������������ " "
    buffSplit() = Split(Trim(StrConv(��������, vbUpperCase)), Chr(32))
    
    ' �������� �������� - �������
    If InStr(1, buffSplit(0), Chr(45), 1) Then
        ' ���� ������� ������� �� 2-� ����, ��
        buffLastName() = Split(Trim(buffSplit(0)), Chr(45))
        j = 0
        i = 0
        Do While j <= UBound(buffLastName())
            buffLastName(j) = StrConv(buffLastName(j), vbProperCase)
            j = j + 1
        Loop
        j = 0
        Do While i <= UBound(buffLastName())
            LastName = LastName + buffLastName(i) + "-"
            i = i + 1
            If i = UBound(buffLastName()) Then
                LastName = LastName + buffLastName(i)
                Exit Do
            End If
        Loop
        i = 0
        Else
            LastName = StrConv(Trim(buffSplit(0)), vbProperCase)
    End If
    
    ' �������� �������� - ���
    If InStr(1, buffSplit(1), Chr(45), 1) Then
        ' ���� ��� ������� �� 2-� ����, ��
        flagFirstName = 1
        buffFirstName() = Split(Trim(buffSplit(1)), Chr(45))
        j = 0
        i = 0
        Do While j <= UBound(buffFirstName())
            buffFirstName(j) = StrConv(buffFirstName(j), vbProperCase)
            j = j + 1
        Loop
        j = 0
        Do While i <= UBound(buffFirstName())
            buffFirstName(i) = StrConv(buffFirstName(i), vbProperCase)
            i = i + 1
        Loop
        i = 0
        Else
            FirstName = StrConv(Trim(buffSplit(1)), vbProperCase)
    End If
    
    ' �������� �������� - ��������
    If InStr(1, buffSplit(2), Chr(45), 1) Then
        ' ���� �������� ������� �� 2-� ����, ��
        flagMidName = 1
        buffMidName() = Split(Trim(buffSplit(2)), Chr(45))
        i = 0
        Do While i <= UBound(buffMidName())
            buffMidName(i) = StrConv(buffMidName(i), vbProperCase)
            i = i + 1
        Loop
        i = 0
        Else
            MidName = StrConv(Trim(buffSplit(2)), vbProperCase)
    End If
    
    ' �������� �����
    If flagFirstName <> 0 Then
        k = 0
        Do While k <= UBound(buffFirstName())
            FirstName = FirstName + Left(buffFirstName(k), 1) + "-"
            k = k + 1
            If k = UBound(buffFirstName()) Then
                FirstName = FirstName + Left(buffFirstName(k), 1) + "."
                Exit Do
            End If
        Loop
        k = 0
    Else
        FirstName = Left(FirstName, 1) + "."
    End If
    
    ' �������� ��������
    If flagMidName <> 0 Then
        m = 0
        Do While m <= UBound(buffMidName())
            MidName = MidName + Left(buffMidName(m), 1) + "-"
            m = m + 1
            If m = UBound(buffMidName()) Then
                MidName = MidName + Left(buffMidName(m), 1) + "."
                Exit Do
            End If
        Loop
        m = 0
    Else
        MidName = Left(MidName, 1) + "."
    End If
        
    Select Case ���
        Case 1
            'If UBound(buffSplit()) = 3 Then
                '�������� = LastName + " " + FirstName + MidName + Left(buffSplit(3), 1) + "."
            'Else
                �������� = LastName + " " + FirstName + MidName
            'End If
        Case 2
            �������� = FirstName + LastName
        Case 3
            �������� = LastName + " " + FirstName + " " + MidName
        Case 4
            �������� = FirstName + MidName + LastName
    End Select

End Function

'Function old_��������(������� As Range, ��� As Range, �������� As Range, ��� As Integer) As Variant
'    Select Case ���
'        Case 1
'            old_�������� = StrConv(�������, vbProperCase) + " " + Left(���, 1) + "." + Left(��������, 1) + "."
'        Case 2
'            old_�������� = Left(���, 1) + "." + StrConv(�������, vbProperCase)
'        Case 3
'           old_�������� = StrConv(�������, vbProperCase) + " " + Left(���, 1) + ". " + Left(��������, 1) + "."
'        Case 4
'            old_�������� = Left(���, 1) + "." + Left(��������, 1) + "." + StrConv(�������, vbProperCase)
'    End Select
'End Function

Function ���������_�������(�������� As Range) As Variant
    '
    ' =����(���(M2="�-�";M2="�/�-�";M2="�-�";M2="�-�";M2="��. �-�";M2="�-�";M2="��. �-�");"�������";
    '       ����(���(M2="��-�";M2="��. ��-�");"����������";
    '       ����(���(M2="��. �-�";M2="�-�";M2="��. �-�";M2="��-��");"��������";
    '       ����(���(M2="���.";M2="���.");"�������";""))))
    '
    If (�������� = "�/�-�") _
        Or (�������� = "�-�") _
        Or (�������� = "�-�") _
        Or (�������� = "��. �-�") _
        Or (�������� = "��. �-� �/�") _
        Or (�������� = "��. �-� (*)") Then
    ���������_������� = "�������"
    ElseIf (�������� = "��. ��-�") _
        Or (�������� = "��. ��-� (*)") _
        Or (�������� = "��-� (*)") _
        Or (�������� = "��-�") Then
    ���������_������� = "����������"
    ElseIf (�������� = "�-� (*)") _
        Or (�������� = "��. �-�") _
        Or (�������� = "��. �-� (*)") Then
    ���������_������� = "��������"
    ElseIf (�������� = "���.") _
        Or (�������� = "���.") _
        Or (�������� = "���. (*)") _
        Or (�������� = "���. (*)") Then
    ���������_������� = "�������"
    Else: ���������_������� = "������"
    End If
End Function
Function ���������_�����������(�������� As Range) As Variant
'Function ���������_�����������(�������� As Range, ��������� As Range) As Variant
    '����� = ���������.Value

    '
    ' =����(���(M2="�-�";M2="�/�-�";M2="�-�";M2="�-�";M2="��. �-�";M2="�-�";M2="��. �-�");"�������";
    '       ����(���(M2="��-�";M2="��. ��-�");"����������";
    '       ����(���(M2="��. �-�";M2="�-�";M2="��. �-�";M2="��-��");"��������";
    '       ����(���(M2="���.";M2="���.");"�������";""))))
    '
    If (�������� = "�-�") _
        Or (�������� = "�/�-�") _
        Or (�������� = "�-� �/�") _
        Or (�������� = "�-�") _
        Or (�������� = "�-�") _
        Or (�������� = "��. �-� �/�") _
        Or (�������� = "��. �-�") _
        Or (�������� = "�-� �/�") _
        Or (�������� = "�-�") _
        Or (�������� = "��. �-�") Then
    ���������_����������� = "�������"
    ElseIf (�������� = "��-��") _
        Or (�������� = "��. �-�") _
        Or (�������� = "�-�") _
        Or (�������� = "��. �-�") Then
    ���������_����������� = "��������"
    ElseIf (�������� = "��. ��-�") _
        Or (�������� = "��-�") Then
    ���������_����������� = "����������"
    ElseIf (�������� = "���.") _
        Or (�������� = "���.") Then
    ���������_����������� = "�������"
    End If
    
    '    If (��������� = "��������") Or _
'        (��������� = "������� ������") Then
'                ���������_����������� = "����������"
End Function

'Function ����������(������ As Range) As Variant
'    Dim CurrColor
'    CurrColor = ������.Cells.Interior.Color
'    ���������� = CurrColor
'End Function

Function ��������������(�������_������ As Range, ���_��������������� As String) As Variant
    Dim Sum As Long
    Dim MaxCells As Long
    
    MaxCells = Application.WorksheetFunction.Max( _
        Application.Caller.Cells.Count, �������_������.Cells.Count)
    ReDim Result(1 To MaxCells, 1 To 1)
    For Each Rng In �������_������.Cells
        If Rng.Value = ���_��������������� Then
            Sum = Sum + 1
        End If
    Next Rng
    
    �������������� = Sum
    
End Function

Function ������������������������������(�������_������_������������� As Range, �������_������_����� As Range, ������������� As String) As Variant
    Dim Sum As Long
    Dim MaxCells As Long
    
    ' MaxCells = Application.WorksheetFunction.Max( _
    '     Application.Caller.Cells.Count, �������_������_�������������.Cells.Count)
    ' ReDim Result(1 To MaxCells, 1 To 1)
'    For Each Rng In �������_������_�������������.Cells
'        If Rng.Value = ������������� Then
'             Sum = Sum + 1
'            For Each Rng2 In �������_������_�����.Cells
'                If Rng2.Value = "�/�" Then
'                    Sum = Sum + 1
'                End If
'                Next Rng2
'        End If
'    Next Rng
    
    For Each Rng In �������_������_�����.Cells
        If Rng.Value = "�/�" Then
            For Each Rng2 In �������_������_�������������.Cells
                If Rng2.Value = ������������� Then
                    Sum = Sum + 1
                End If
                Next Rng2
        End If
    Next Rng
    
    ������������������������������ = Sum
    
End Function

Function ���������(�������_������ As Range) As Variant
    Dim N As Long
    Dim N2 As Long
    Dim Rng As Range
    Dim MaxCells As Long
    Dim Result() As Variant
    Dim R As Long
    Dim C As Long
    
    MaxCells = Application.WorksheetFunction.Max( _
        Application.Caller.Cells.Count, �������_������.Cells.Count)
    ReDim Result(1 To MaxCells, 1 To 1)
    For Each Rng In �������_������.Cells
        If Rng.Value <> vbNullString Then
            N = N + 1
            Result(N, 1) = Rng.Value
        End If
    Next Rng
    For N2 = N + 1 To MaxCells
        Result(N2, 1) = vbNullString
    Next N2
    
    If Application.Caller.Rows.Count = 1 Then
        ��������� = Application.WorksheetFunction.Transpose(Result)
    Else
        ��������� = Result
    End If
End Function

Public Function ������(���� As Integer, ������������ As Integer) As String
    Select Case ������������
    Case 1
        Select Case ����
        Case 1
            ������ = "�����������"
        Case 2
            ������ = "�������"
        Case 3
            ������ = "�����"
        Case 4
            ������ = "�������"
        Case 5
            ������ = "�������"
        Case 6
            ������ = "�������"
        Case 7
            ������ = "�����������"
        Case Else
            ������ = "������������ �������� ��� ��� ������ [1..7]"
        End Select
            Case Else
        ������ = "������������ �������� ��� ������ ������ [1..12]"
    End Select
End Function
        
Public Function �������(����� As Integer, ����� As Integer) As String

    Dim arrMonth(1 To 12, 1 To 6) As Variant
    '
    ' ���������� �������
    ' ���� (�,�), ���
    '           � - �����
    '           y - �����
    '
    arrMonth(1, 1) = "������"
    arrMonth(1, 2) = "������"
    arrMonth(1, 3) = "������"
    arrMonth(1, 4) = "������"
    arrMonth(1, 5) = "�������"
    arrMonth(1, 6) = "������"
    
    arrMonth(2, 1) = "�������"
    arrMonth(2, 2) = "�������"
    arrMonth(2, 3) = "�������"
    arrMonth(2, 4) = "�������"
    arrMonth(2, 5) = "��������"
    arrMonth(2, 6) = "�������"
    
    arrMonth(3, 1) = "����"
    arrMonth(3, 2) = "�����"
    arrMonth(3, 3) = "�����"
    arrMonth(3, 4) = "����"
    arrMonth(3, 5) = "������"
    arrMonth(3, 6) = "�����"
    
    arrMonth(4, 1) = "������"
    arrMonth(4, 2) = "������"
    arrMonth(4, 3) = "������"
    arrMonth(4, 4) = "������"
    arrMonth(4, 5) = "�������"
    arrMonth(4, 6) = "������"
    
    arrMonth(5, 1) = "���"
    arrMonth(5, 2) = "���"
    arrMonth(5, 3) = "���"
    arrMonth(5, 4) = "���"
    arrMonth(5, 5) = "����"
    arrMonth(5, 6) = "���"
    
    arrMonth(6, 1) = "����"
    arrMonth(6, 2) = "����"
    arrMonth(6, 3) = "����"
    arrMonth(6, 4) = "����"
    arrMonth(6, 5) = "�����"
    arrMonth(6, 6) = "����"
    
    arrMonth(7, 1) = "����"
    arrMonth(7, 2) = "����"
    arrMonth(7, 3) = "����"
    arrMonth(7, 4) = "����"
    arrMonth(7, 5) = "�����"
    arrMonth(7, 6) = "����"
    
    arrMonth(8, 1) = "������"
    arrMonth(8, 2) = "�������"
    arrMonth(8, 3) = "�������"
    arrMonth(8, 4) = "������"
    arrMonth(8, 5) = "��������"
    arrMonth(8, 6) = "�������"
    
    arrMonth(9, 1) = "��������"
    arrMonth(9, 2) = "��������"
    arrMonth(9, 3) = "��������"
    arrMonth(9, 4) = "��������"
    arrMonth(9, 5) = "���������"
    arrMonth(9, 6) = "��������"
    
    arrMonth(10, 1) = "�������"
    arrMonth(10, 2) = "�������"
    arrMonth(10, 3) = "�������"
    arrMonth(10, 4) = "�������"
    arrMonth(10, 5) = "��������"
    arrMonth(10, 6) = "�������"
    
    arrMonth(11, 1) = "������"
    arrMonth(11, 2) = "������"
    arrMonth(11, 3) = "������"
    arrMonth(11, 4) = "������"
    arrMonth(11, 5) = "�������"
    arrMonth(11, 6) = "������"
   
    arrMonth(12, 1) = "�������"
    arrMonth(12, 2) = "�������"
    arrMonth(12, 3) = "�������"
    arrMonth(12, 4) = "�������"
    arrMonth(12, 5) = "��������"
    arrMonth(12, 6) = "�������"
    
    ������� = arrMonth(�����, �����)
    
End Function

Function ������_����(����� As Variant, ��� As Integer) As String
    
    Dim sDay As String
    Dim sMonth As String
    Dim sYear As String
        
    ����� = StrConv(Trim(�����), vbUpperCase)
    
    Select Case �����
        Case "�������"
            dDate = 0
        Case "������"
            dDate = 1
        Case Else:
            If Conversion.Int(�����) >= 0 Then
                dDate = Conversion.Int(�����)
            End If
    End Select
    
    sDay = CStr(DatePart("d", DateAdd("d", dDate, Now), vbMonday))
    sMonth = StrConv(CStr(�������(DatePart("m", DateAdd("d", dDate, Now), vbMonday), 2)), vbLowerCase)
    sYear = CStr(DatePart("yyyy", DateAdd("d", dDate, Now), vbMonday))
    
    If (����� = "�������") Then
        Select Case ���
            Case 1
                ������_���� = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "����"
            Case 2
                ������_���� = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "�."
            Case 3
                ������_���� = "�" + sDay + "�" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "����"
            Case 4
                ������_���� = "�" + sDay + "�" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "�."
        End Select
        ElseIf (����� = "������") Then
            Select Case ���
                Case 1
                    ������_���� = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "����"
                Case 2
                    ������_���� = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "�."
                Case 3
                    ������_���� = "�" + sDay + "�" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "����"
                Case 4
                    ������_���� = "�" + sDay + "�" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "�."
            End Select
            Else:
                Select Case ���
                    Case 1
                        ������_���� = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "����"
                    Case 2
                        ������_���� = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "�."
                    Case 3
                        ������_���� = "�" + sDay + "�" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "����"
                    Case 4
                        ������_���� = "�" + sDay + "�" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "�."
                End Select
    End If
    
End Function

Public Function �������(���������_���� As Date, ��������_���� As Date) As String
    Dim ���� As Integer
    Dim ������  As Integer
    Dim �����  As Integer
    
    Dim �������� As String
    Dim ������� As String
    Dim ��������� As String
    
    ���� = WorksheetFunction.RoundDown(((��������_���� - ���������_����) / 365), 0)
    ������ = WorksheetFunction.RoundDown(((((��������_���� - ���������_����) / 365) - ����) * 12), 0)
    ����� = WorksheetFunction.RoundDown(((((((��������_���� - ���������_����) / 365) - ����) * 12) - ������) * (365 / 12)), 0)
    
    If (���� = 1) Then
        �������� = " ���"
        ElseIf (���� >= 2) And (���� < 5) Then
        �������� = " ����"
            ElseIf (���� >= 5) Or (���� = 0) Then
            �������� = " ���"
    End If
    
    If (������ = 1) Then
        ������� = " �����"
        ElseIf (������ >= 2) And (������ < 5) Then
        ������� = " ������"
            ElseIf (������ >= 5) Or (������ = 0) Then
            ������� = " �������"
    End If
    
    If (����� = 1) Then
        ��������� = " ����"
        ElseIf (����� >= 2) And (����� < 5) Then
        ��������� = " ���"
            ElseIf (����� >= 5) Or (����� = 0) Then
            ��������� = " ����"
    End If

    If ���� = 0 Then
        ������� = CStr(������) + ������� + " " + CStr(�����) + ���������
        ElseIf ������ = 0 And ����� = 0 Then
            ������� = CStr(����) + ��������
            ElseIf ������ = 0 Then
            ������� = CStr(����) + �������� + " " + CStr(�����) + ���������
                ElseIf ����� = 0 Then
                ������� = CStr(����) + �������� + " " + CStr(������) + �������
                Else
                ������� = CStr(����) + �������� + " " + CStr(������) + ������� + " " + CStr(�����) + ���������
    End If
End Function

Public Function �����������(�������_������ As Range) As Variant
    Dim ���� As Integer
    Dim ������  As Integer
    Dim �����  As Integer
    
    Dim �������� As Range
    ' �������()
    ' WorksheetFunction.Quotient
    '
    ' ����� ()
    ' WorksheetFunction.Mod()
    
    For Each �������� In �������_������.Cells
        ���� = ���� + WorksheetFunction.Quotient(������, 12)
        ������ = WorksheetFunction.Mod(������, 12) + WorksheetFunction.Quotient(�����, 30)
        ����� = WorksheetFunction.Mod(�����, 30)
    Next ��������
   
End Function

Public Function ��������������������������������(������������� As String) As String
    '
    ' 1 ��� (�)
    '
    ������������� = LCase(�������������)
    
    If ������������� = "���������� 1���" Then
        �������������������������������� = "����������"
        
    ElseIf ������������� = "1 ���" Then
        �������������������������������� = "1 ���"
        
    ElseIf ������������� = "2 ���" Then
        �������������������������������� = "2 ���"
        
    ElseIf ������������� = "3 ���" Then
        �������������������������������� = "3 ���"
        
    ElseIf ������������� = "���������� �������" Then
        �������������������������������� = "�������"
        
    ElseIf ������������� = "1 ���������� �����" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "1 ������ �����" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "1 ����� �����" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "1 ���.�������." Then
        �������������������������������� = "���"
        
    ElseIf ������������� = "1 ��" Then
        �������������������������������� = "��������"
        
    ElseIf ������������� = "���������� 2���" Then
        �������������������������������� = "����������"
        
    ElseIf ������������� = "4 ���" Then
        �������������������������������� = "4 ���"
        
    ElseIf ������������� = "5 ���" Then
        �������������������������������� = "5 ���"
        
    ElseIf ������������� = "6 ���" Then
        �������������������������������� = "6 ���"
        
    ElseIf ������������� = "2 ���������� �������" Then
        �������������������������������� = "�������"
        
    ElseIf ������������� = "2 ���������� �����" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "2 ������ �����" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "2 ����� �����" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "2 ���.�������." Then
        �������������������������������� = "���"
        
    ElseIf ������������� = "2 ��" Then
        �������������������������������� = "��������"
    '
    ' � ������������
    '
    ElseIf (������������� = "������.��.") Or _
            (������������� = "������.��-�.") Or _
            (������������� = "������.��-�") Or _
            (������������� = "������.����.") Or _
            (������������� = "������.����.") Or _
            (������������� = "������.����.") Or _
            (������������� = "������.����") Then
        �������������������������������� = "����"
        
    '
    ' ��(�)
    '
    ElseIf ������������� = "��(�)" Then
        �������������������������������� = "��(�)"
        
    '
    ' �����
    '
    ElseIf ������������� = "���������� �����" Then
        �������������������������������� = "���. �����"
        
    ElseIf ������������� = "1 �������" Then
        �������������������������������� = "1 �������"
    
    ElseIf ������������� = "2 �������" Then
        �������������������������������� = "2 �������"
        
    ElseIf ������������� = "3 �������" Then
        �������������������������������� = "3 �������"
        
    ElseIf ������������� = "���" Then
        �������������������������������� = "���"
        
    ElseIf ������������� = "���" Then
        �������������������������������� = "���"
    
    '
    ' ������������� ��������
    '
    ElseIf ������������� = "�������" Then
        �������������������������������� = "�������"
    
    ElseIf ������������� = "1 ������ ����" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "3 ������ ���� (����)" Then
        �������������������������������� = "����"
        
    ElseIf ������������� = "���" Then
        �������������������������������� = "���"
    
    ElseIf ������������� = "������� ���" Then
        �������������������������������� = "���"
    
    ElseIf ������������� = "����� ���" Then
        �������������������������������� = "���"
       
    ElseIf ������������� = "���.�." Then
        �������������������������������� = "���.����"
        
    ElseIf ������������� = "��" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "����" Then
        �������������������������������� = "����"
        
    ElseIf ������������� = "���.�����" Then
        �������������������������������� = "���. ������"
    
    ElseIf ������������� = "���. ������" Then
        �������������������������������� = "���. ������"
        
    '
    ' ��������� �������������
    '
    ElseIf ������������� = "��" Then
        �������������������������������� = "��"
        
    ElseIf ������������� = "�������" Then
        �������������������������������� = "�������"
        
    ElseIf ������������� = "�����" Then
        �������������������������������� = "�����"
        
    ElseIf ������������� = "�����" Then
        �������������������������������� = "�����"
        
    ElseIf ������������� = "�� �����" Then
        �������������������������������� = "�� �����"
        
    ElseIf ������������� = "�� �����" Then
        �������������������������������� = "�� �����"
        
    ElseIf ������������� = "���" Then
        �������������������������������� = "����"
        
    ElseIf ������������� = "������ ���" Then
        �������������������������������� = "����"
        
    ElseIf ������������� = "���" Then
        �������������������������������� = "���"
        
    ElseIf ������������� = "���" Then
        �������������������������������� = "���"
        
    '
    ' ��� ������������� �� ������������� (��� �������)
    '
    Else: �������������������������������� = "������"
    End If
End Function

Private Function ���������_���(������������ As String, ����� As Integer) As String
    ��������� = LCase(Mid(CStr(������������), (Len(CStr(������������)) - 1), Len(CStr(������������))))
    ���������_����������� = LCase(Mid(CStr(������������), (Len(CStr(������������)) - 2), Len(CStr(������������))))
    Select Case �����
        ' ������������
        Case 1
            ���������_��� = ������������
            
        ' �����������
        Case 2
            ' �������, �������
            If (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
                    
            ' �����
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
                    
            ' ������
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
                    
            ' �����, ������
            ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
                    
            ' ����
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
            
            ' ������
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
                
            ' �������
            ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
                    
            '
            ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = ������������ + "�"
             ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = ������������ + "�"
             ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = ������������ + "�"
            ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = ������������
            Else: ���������_��� = "!!!"
            End If
            
        ' ���������
        Case 3
            ' �������, �������
            If (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
                    
            ' �����
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
                    
            ' ������
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
                    
            ' �����, ������
            ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
                    
            ' ����
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
            
            ' ������
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
                
            ' �������
            ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
                
            ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = ������������ + "�"
             
             '
             ' ����������� ��������� 3-� ���������
             '
             ' ������
             ElseIf (���������_����������� = "���") Then
                    ���������_��� = ������������ + "�"
            ' ������
             ElseIf (���������_����������� = "���") Then
                    ���������_��� = (Mid(������������, 1, Len(������������) - 3)) + "��"
                    
             ' �����
             ElseIf (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Then
                    ���������_��� = (Mid(������������, 1, Len(������������) - 1)) + "�"
                    
             ' �������, �������
             ElseIf (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Then
                    ���������_��� = ������������ + "�"
             
             ' �������
             ElseIf (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Or _
                    (���������_����������� = "���") Then
                    ���������_��� = ������������
             
             ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = ������������ + "�"
             ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = ������������ + "�"
            ElseIf (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Or _
                    (��������� = "��") Then
                ���������_��� = ������������
            
            ' ������� �����
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"            '
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
            ElseIf (��������� = "��") Then
                ���������_��� = (Mid(������������, 1, Len(������������) - 2)) + "��"
                
            Else: ���������_��� = "!!!"
            End If
                       
        ' �����������
        Case 4
            '���������_��� =
            
        ' ������������
        Case 5
            '���������_��� =
        
        ' ����������
        Case 6
            '���������_��� =
        Case Else
            ���������_��� = "������������ �������� ��� ������ [1..6]"
    End Select
End Function

Private Function ���������_��������(������������ As String, ����� As Integer) As String
    ��������� = LCase(Mid(CStr(������������), (Len(CStr(������������)) - 2), Len(CStr(������������))))
    Select Case �����
        ' ������������
        Case 1
            ���������_�������� = ������������
            
        ' �����������
        Case 2
            If (��������� = "���") Then
                ���������_�������� = ������������ + "�"
            ElseIf (��������� = "���") Or _
                    (��������� = "���") Or _
                    (��������� = "���") Or _
                    (��������� = "���") Or _
                    (��������� = "���") Then
                ���������_�������� = ������������
            End If
        ' ���������
        Case 3
            If (��������� = "���") Then
                ���������_�������� = ������������ + "�"
            ElseIf (��������� = "���") Or _
                    (��������� = "���") Or _
                    (��������� = "���") Or _
                    (��������� = "���") Or _
                    (��������� = "���") Or _
                    (��������� = "���") Or _
                    (��������� = "���") Then
                ���������_�������� = ������������
            ' ������� ��������
            ElseIf (��������� = "���") Then
                ���������_�������� = (Mid(������������, 1, Len(������������) - 2)) + "��"
            End If
            
        ' �����������
        Case 4
            '���������_�������� =
            
        '
        Case 5
            '���������_�������� =
            
        '
        Case 6
            '���������_�������� =
            
        '
        Case Else
            ���������_�������� = "������������ �������� ��� ������ [1..6]"
    End Select
End Function

Private Function ���������_�������(������������ As String, ����� As Integer) As String
    If ((Len(CStr(������������)) - 1) > 0) Then
        ��������� = LCase(Mid(CStr(������������), (Len(CStr(������������)) - 1), Len(CStr(������������))))
        ���������_����������� = LCase(Mid(CStr(������������), (Len(CStr(������������)) - 2), Len(CStr(������������))))
    Else:
        ���������_������� = ������������
    End If
    
    Select Case �����
        ' ������������
        Case 1
            ���������_������� = ������������
            
        ' �����������
        Case 2
            If (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase(������������ + "�")
            ElseIf (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 2)) + "���")
            ElseIf (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 1)) + "�")
            ElseIf (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 1)) + "�")
            ElseIf (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase(������������)
                
            ' ����������� ���������
            ElseIf (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Then
                ���������_������� = UCase(������������)
            ' ������, ����
            ElseIf (���������_����������� = "���") Or _
                   (���������_����������� = "���") Then
                ���������_������� = UCase(������������)
                
            Else: ���������_������� = "!!!"
            End If
        ' ���������
        Case 3
            If (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase(������������ + "�")
            ElseIf (��������� = "��") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 2)) + "�")
               
            ElseIf (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 2)) + "���")
                
            ElseIf (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 1)))
            ElseIf (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Or _
                (��������� = "��") Then
                ���������_������� = UCase(������������)
            
            ' ������� �������
            ElseIf (��������� = "��") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 1)) + "��")
                
            ' ����������� ���������
            ElseIf (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Then
                ���������_������� = UCase(������������)
            ' ������, ����
            ElseIf (���������_����������� = "���") Or _
                   (���������_����������� = "���") Then
                ���������_������� = UCase(������������)
            
            ' �������
            ElseIf (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Or _
                   (���������_����������� = "���") Then
                ���������_������� = UCase(������������ + "�")
            
            ' �������
            ElseIf (���������_����������� = "���") Then
                ���������_������� = UCase(������������ + "�")
            
            ' ������, ��������
            ElseIf (���������_����������� = "���") Or _
                   (���������_����������� = "���") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 1)) + "�")
            
            ' ���������
            ElseIf (���������_����������� = "���") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 2)) + "��")
            
            ' ��������
            ElseIf (���������_����������� = "���") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 1)) + "��")
                
            ' �������
            ElseIf (���������_����������� = "���") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 1)) + "�")
            
            ' ������
            ElseIf (���������_����������� = "���") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 1)) + "�")
                
            ' ������� �������
            ElseIf (���������_����������� = "���") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 2)) + "��")
            ElseIf (���������_����������� = "���") Then
                ���������_������� = UCase((Mid(������������, 1, Len(������������) - 2)) + "��")
                
                Else: ���������_������� = "!!!"
            End If
       
        ' �����������
        Case 4
            ���������_������� = ���������
        
        ' ������������
        Case 5
            '���������_������� =
        
        ' ����������
        Case 6
            '���������_������� =
        Case Else
            ���������_������� = "������������ �������� ��� ������ [1..6]"
    End Select
End Function

Function ����������(�������� As String) As Variant
    
    Dim buffLastName As String
    Dim buffFirstName As String
    Dim buffMidName As String
    Dim buffAppendName As String
    
    Dim buffSplit() As String
    
    Dim t As Integer
 
    ' ������ ������� ����� � ������
    �������� = Trim(��������)
    ' ������ ������� �������
    Do While InStr(1, ��������, Space(2), 1) <> 0
        �������� = Replace(��������, Space(2), Space(1), vbTextCompare)
    Loop
    
    ' ������� ������ � ������������ " "
    buffSplit() = Split(Trim(��������), Chr(32))
    
    ' �������� �������� - �������
    buffLastName = buffSplit(0)
    ' �������� �������� - ���
    buffFirstName = buffSplit(1)
    ' �������� �������� - ��������
    buffMidName = buffSplit(2)
    
    t = 1
    
    On Error GoTo C
        buffAppendName = buffSplit(3)
        t = 2
C:
    
    Select Case t
        Case 1
            ���������� = StrConv(buffLastName, vbUpperCase) + " " + StrConv(buffFirstName, vbProperCase) + " " + StrConv(buffMidName, vbProperCase)
            t = 0
        Case 2
            ���������� = StrConv(buffLastName, vbUpperCase) + " " + StrConv(buffFirstName, vbProperCase) + " " + StrConv(buffMidName, vbProperCase) + " " + StrConv(buffAppendName, vbProperCase)
            t = 0
    End Select
End Function

Function ������_��������_������(������ As String) As String
    If ������ = "�-�" Then
        ������_��������_������ = "���������"
        ElseIf ������ = "�/�-�" Then
            ������_��������_������ = "������������"
        ElseIf ������ = "�-�" Then
            ������_��������_������ = "�����"
        ElseIf ������ = "�-�" Then
            ������_��������_������ = "�������"
        ElseIf ������ = "��. �-�" Then
            ������_��������_������ = "������� ���������"
        ElseIf ������ = "�-�" Then
            ������_��������_������ = "���������"
        ElseIf ������ = "��. �-�" Then
            ������_��������_������ = "������� ���������"
        ElseIf ������ = "��. ��-�" Then
            ������_��������_������ = "������� ���������"
        ElseIf ������ = "��-�" Then
            ������_��������_������ = "���������"
        ElseIf ������ = "��-��" Then
            ������_��������_������ = "��������"
        ElseIf ������ = "��. �-�" Then
            ������_��������_������ = "������� �������"
        ElseIf ������ = "�-�" Then
            ������_��������_������ = "�������"
        ElseIf ������ = "��. �-�" Then
            ������_��������_������ = "������� �������"
        ElseIf ������ = "���." Then
            ������_��������_������ = "��������"
        ElseIf ������ = "���." Then
            ������_��������_������ = "�������"
        Else:
            ������_��������_������ = "!!!"
    End If
End Function

Function �������������(���� As String, ����� As String, ��������� As String) As String

    Dim ��� As Integer
    ��� = 0
    
    If (���� = "1 ���") Or _
       (���� = "2 ���") Or _
       (���� = "3 ���") Or _
       (���� = "���������� �������") Then
            ��� = 1
        ElseIf (���� = "1 ���������� �����") Or _
            (���� = "1 ������ �����") Or _
            (���� = "1 ����� �����") Or _
            (���� = "1 ���.�������.") Or _
            (���� = "1 ��") Then
                ��� = 2
    End If
    
    Select Case ���
        Case 0
            If ���� = "���������� 1���" Then
                    ������������� = "���. 1 ���(�)"
                ElseIf (���� = "���������� 1���") And (����� = "����") Then
                    ������������� = "���. 1 ���(�)"
            End If
        Case 1
            ������������� = ����
        Case 2
            ������������� = �����
    End Select
End Function



