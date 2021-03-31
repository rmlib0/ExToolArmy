Attribute VB_Name = "functions"
' --------------------------------------------------------------------------------
'
'       Title:      ExToolArmy.xlam
'
'       Purpose:    Набор функций, для внутреннего пользования в среде
'                   Microsoft Office 2017 Excel
'
'       License:    BearLic v.0.1a
'
'       DateTime:   05/14/2020 14:00 GMT+3
'
'       Author:     rmlib.null (OVSYANNIKOV Mikhail Mikhailovich)
'                              (ОВСЯННИКОВ Михаил Михайлович)
'
'       Contacts:   rmlib.null@gmail.com
'                   +7 (918) 518 40-62
'
' --------------------------------------------------------------------------------

Public Function ЦветЯчейки(Ячейка As Range) As Double
    ЦветЯчейки = Ячейка.Interior.Color
End Function

Public Function Склонение(Значение As String, Падеж As Integer) As Variant
    Dim buffLastName As String
    Dim buffFirstName As String
    Dim buffMidName As String
    Dim buffPostMidName As String
    Dim buffSplit() As String
        
    Dim gradLastName As String
    Dim gradFirstName As String
    Dim gradMidName As String
    
    ' Убрать двойные пробелы во входной строке
    Do While InStr(1, Значение, Space(2), 1) <> 0
        Значение = Replace(Значение, Space(2), Space(1), vbTextCompare)
    Loop

    ' Разбить строку с разделителем " "
    buffSplit() = Split(Trim(Значение), Chr(32))
    
    ' Получить значение - Фамилия
    buffLastName = buffSplit(0)
    ' Получить значение - Имя
    buffFirstName = buffSplit(1)
    ' Получить значение - Отчество
    buffMidName = buffSplit(2)
    ' TODO
    ' Получить значение - Пост отчество типа Оглы
    ' buffPostMidName = buffSplit(3)
        
    gradLastName = СКЛОНЕНИЕ_ФАМИЛИЯ(buffLastName, Падеж)
    gradFirstName = СКЛОНЕНИЕ_ИМЯ(buffFirstName, Падеж)
    gradMidName = СКЛОНЕНИЕ_ОТЧЕСТВО(buffMidName, Падеж)
    
    Склонение = gradLastName & Chr(32) & gradFirstName & Chr(32) & gradMidName & Chr(32) & buffPostMidName

End Function

Public Function СклонениеЗвания(Значение As String, Падеж As Integer) As Variant
    Dim buffSplit() As String
    Dim buffSimpleGrade As String
    Dim buffComplexGrade As String
    Dim buffOtherGrade As String
    
    Dim flagSimpleGrade As Integer
    Dim flagComplexGrade As Integer
    Dim flagOtherGrade As Integer
         
    Значение = StrPrepare(Значение)
    
    ' Убрать пробелы слева и справа
    'Значение = Trim(Значение)
    
    ' Убрать двойные пробелы
    'Do While InStr(1, Значение, Space(2), 1) <> 0
    '    Значение = Replace(Значение, Space(2), Space(1), vbTextCompare)
    'Loop
    
    ' Разбить строку с разделителем " "
    buffSplit() = Split(Значение, Chr(32))
        
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
        Select Case Падеж
            ' Именительный
            Case 1
                If flagSimpleGrade = 1 Then
                    buffSimpleGrade = buffSplit(0)
                    СклонениеЗвания = buffSimpleGrade
                    If flagComplexGrade = 1 Then
                        buffSimpleGrade = buffSplit(0)
                        buffComplexGrade = buffSplit(1)
                        СклонениеЗвания = buffSimpleGrade + Chr(32) + buffComplexGrade
                        If flagOtherGrade = 1 Then
                            buffSimpleGrade = buffSplit(0)
                            buffComplexGrade = buffSplit(1)
                            buffOtherGrade = buffSplit(2)
                            СклонениеЗвания = buffSimpleGrade + Chr(32) + buffComplexGrade + Chr(32) + buffOtherGrade
                        End If
                    End If
                End If
                
            ' Родительный
            Case 2
                If flagSimpleGrade = 1 Then
                        ' полковник
                        ' подполковник
                    If (extEndingSimpleGrade = "ник") Then
                        СклонениеЗвания = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "ика"
                        ' майор
                        ElseIf (extEndingSimpleGrade = "йор") Then
                            СклонениеЗвания = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "ора"
                        ' капитан
                        ElseIf (extEndingSimpleGrade = "тан") Then
                            СклонениеЗвания = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "ана"
                        ' лейтенант
                        ' сержант
                        ElseIf (extEndingSimpleGrade = "ант") Then
                            СклонениеЗвания = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "нта"
                        ' прапорщик
                        ElseIf (extEndingSimpleGrade = "щик") Then
                            СклонениеЗвания = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "ика"
                        ' старшина
                       ElseIf (extEndingSimpleGrade = "ина") Then
                            СклонениеЗвания = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "ны"
                        ' ефрейтор
                        ElseIf (extEndingSimpleGrade = "тор") Then
                            СклонениеЗвания = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "ора"
                        ' рядовой
                        ElseIf (extEndingSimpleGrade = "вой") Then
                            СклонениеЗвания = (Mid(buffSimpleGrade, 1, Len(buffSimpleGrade) - 2)) + "ого"
                        Else: СклонениеЗвания = "!!!"
                    End If
                End If
        End Select
'    End If
    
    'If InStr(1, buffSplit(1), Chr(45), 1) Then
        ' Если ИМЯ состоит из 2-х слов, то
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
'            СклонениеЗвания = StrConv(Trim(buffSplit(1)), vbProperCase)
'    End If
    
    ' Получить значение - звание простое
    'buffSimpleGrade = buffSplit(0)
    ' Получить значение - составное звание
    'buffComplexGrade = buffSplit(1)
    ' Получить значение - мед., юрист и прочее
    'buffOtherGrade = buffSplit(2)

End Function


Private Function ПереводИнформации(Значение As String) As String
    Dim ArrayCyrillic(1 To 33) As String
    Dim ArrayLatin(1 To 26) As String
    Dim i As Long
    
    ArrayCyrillic(1) = "а"
    ArrayCyrillic(2) = "б"
    ArrayCyrillic(3) = "в"
    ArrayCyrillic(4) = "г"
    ArrayCyrillic(5) = "д"
    ArrayCyrillic(6) = "е"
    ArrayCyrillic(7) = "ё"
    ArrayCyrillic(8) = "ж"
    ArrayCyrillic(9) = "з"
    ArrayCyrillic(10) = "и"
    ArrayCyrillic(11) = "й"
    ArrayCyrillic(12) = "к"
    ArrayCyrillic(13) = "л"
    ArrayCyrillic(14) = "м"
    ArrayCyrillic(15) = "н"
    ArrayCyrillic(16) = "о"
    ArrayCyrillic(17) = "п"
    ArrayCyrillic(18) = "р"
    ArrayCyrillic(19) = "с"
    ArrayCyrillic(20) = "т"
    ArrayCyrillic(21) = "у"
    ArrayCyrillic(22) = "ф"
    ArrayCyrillic(23) = "х"
    ArrayCyrillic(24) = "ц"
    ArrayCyrillic(25) = "ч"
    ArrayCyrillic(26) = "ш"
    ArrayCyrillic(27) = "щ"
    ArrayCyrillic(28) = "ъ"
    ArrayCyrillic(29) = "ы"
    ArrayCyrillic(30) = "ь"
    ArrayCyrillic(31) = "э"
    ArrayCyrillic(32) = "ю"
    ArrayCyrillic(33) = "я"
    
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

    
    If Len(Значение) = 0 Then
        ПереводИнформации = ""
    Else
        For i = 1 To Len(Значение)
            CurrentChar = Mid(Значение, i, 1)
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
        ПереводИнформации = Trim(ConvertChar)
    End If
End Function

Function ИНИЦИАЛЫ(Значение As String, Тип As Integer) As Variant
    
    Dim buffLastName() As String
    Dim buffFirstName() As String
    Dim buffMidName() As String
    
    Dim buffSplit() As String
       
    Dim flagFirstName As Integer
    Dim flagMidName As Integer
    
    Dim LastName As String
    Dim FirstName As String
    Dim MidName As String
    
    ' Убрать пробелы слева и справа
    Значение = Trim(Значение)
    
    ' Убрать двойные пробелы
    Do While InStr(1, Значение, Space(2), 1) <> 0
        Значение = Replace(Значение, Space(2), Space(1), vbTextCompare)
    Loop
    
    ' Разбить строку с разделителем " "
    buffSplit() = Split(Trim(StrConv(Значение, vbUpperCase)), Chr(32))
    
    ' Получить значение - ФАМИЛИЯ
    If InStr(1, buffSplit(0), Chr(45), 1) Then
        ' Если ФАМИЛИЯ состоит из 2-х слов, то
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
    
    ' Получить значение - ИМЯ
    If InStr(1, buffSplit(1), Chr(45), 1) Then
        ' Если ИМЯ состоит из 2-х слов, то
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
    
    ' Получить значение - ОТЧЕСТВО
    If InStr(1, buffSplit(2), Chr(45), 1) Then
        ' Если ОТЧЕСТВО состоит из 2-х слов, то
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
    
    ' Инициалы ИМЕНИ
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
    
    ' Инициалы ОТЧЕСТВА
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
        
    Select Case Тип
        Case 1
            'If UBound(buffSplit()) = 3 Then
                'ИНИЦИАЛЫ = LastName + " " + FirstName + MidName + Left(buffSplit(3), 1) + "."
            'Else
                ИНИЦИАЛЫ = LastName + " " + FirstName + MidName
            'End If
        Case 2
            ИНИЦИАЛЫ = FirstName + LastName
        Case 3
            ИНИЦИАЛЫ = LastName + " " + FirstName + " " + MidName
        Case 4
            ИНИЦИАЛЫ = FirstName + MidName + LastName
    End Select

End Function

'Function old_ИНИЦИАЛЫ(Фамилия As Range, Имя As Range, Отчество As Range, Тип As Integer) As Variant
'    Select Case Тип
'        Case 1
'            old_ИНИЦИАЛЫ = StrConv(Фамилия, vbProperCase) + " " + Left(Имя, 1) + "." + Left(Отчество, 1) + "."
'        Case 2
'            old_ИНИЦИАЛЫ = Left(Имя, 1) + "." + StrConv(Фамилия, vbProperCase)
'        Case 3
'           old_ИНИЦИАЛЫ = StrConv(Фамилия, vbProperCase) + " " + Left(Имя, 1) + ". " + Left(Отчество, 1) + "."
'        Case 4
'            old_ИНИЦИАЛЫ = Left(Имя, 1) + "." + Left(Отчество, 1) + "." + StrConv(Фамилия, vbProperCase)
'    End Select
'End Function

Function КАТЕГОРИЯ_ШТАТНАЯ(Значение As Range) As Variant
    '
    ' =ЕСЛИ(ИЛИ(M2="п-к";M2="п/п-к";M2="м-р";M2="к-н";M2="ст. л-т";M2="л-т";M2="мл. л-т");"Офицеры";
    '       ЕСЛИ(ИЛИ(M2="пр-к";M2="ст. пр-к");"Прапорщики";
    '       ЕСЛИ(ИЛИ(M2="мл. с-т";M2="с-т";M2="ст. с-т";M2="ст-на");"Сержанты";
    '       ЕСЛИ(ИЛИ(M2="ряд.";M2="ефр.");"Солдаты";""))))
    '
    If (Значение = "п/п-к") _
        Or (Значение = "м-р") _
        Or (Значение = "к-н") _
        Or (Значение = "ст. л-т") _
        Or (Значение = "ст. л-т м/с") _
        Or (Значение = "ст. л-т (*)") Then
    КАТЕГОРИЯ_ШТАТНАЯ = "Офицеры"
    ElseIf (Значение = "ст. пр-к") _
        Or (Значение = "ст. пр-к (*)") _
        Or (Значение = "пр-к (*)") _
        Or (Значение = "пр-к") Then
    КАТЕГОРИЯ_ШТАТНАЯ = "Прапорщики"
    ElseIf (Значение = "с-т (*)") _
        Or (Значение = "ст. с-т") _
        Or (Значение = "ст. с-т (*)") Then
    КАТЕГОРИЯ_ШТАТНАЯ = "Сержанты"
    ElseIf (Значение = "ряд.") _
        Or (Значение = "ефр.") _
        Or (Значение = "ряд. (*)") _
        Or (Значение = "ефр. (*)") Then
    КАТЕГОРИЯ_ШТАТНАЯ = "Солдаты"
    Else: КАТЕГОРИЯ_ШТАТНАЯ = "ОШИБКА"
    End If
End Function
Function КАТЕГОРИЯ_ФАКТИЧЕСКАЯ(Значение As Range) As Variant
'Function КАТЕГОРИЯ_ФАКТИЧЕСКАЯ(Значение As Range, Должность As Range) As Variant
    'Буфер = Должность.Value

    '
    ' =ЕСЛИ(ИЛИ(M2="п-к";M2="п/п-к";M2="м-р";M2="к-н";M2="ст. л-т";M2="л-т";M2="мл. л-т");"Офицеры";
    '       ЕСЛИ(ИЛИ(M2="пр-к";M2="ст. пр-к");"Прапорщики";
    '       ЕСЛИ(ИЛИ(M2="мл. с-т";M2="с-т";M2="ст. с-т";M2="ст-на");"Сержанты";
    '       ЕСЛИ(ИЛИ(M2="ряд.";M2="ефр.");"Солдаты";""))))
    '
    If (Значение = "п-к") _
        Or (Значение = "п/п-к") _
        Or (Значение = "м-р м/с") _
        Or (Значение = "м-р") _
        Or (Значение = "к-н") _
        Or (Значение = "ст. л-т м/с") _
        Or (Значение = "ст. л-т") _
        Or (Значение = "л-т м/с") _
        Or (Значение = "л-т") _
        Or (Значение = "мл. л-т") Then
    КАТЕГОРИЯ_ФАКТИЧЕСКАЯ = "Офицеры"
    ElseIf (Значение = "ст-на") _
        Or (Значение = "ст. с-т") _
        Or (Значение = "с-т") _
        Or (Значение = "мл. с-т") Then
    КАТЕГОРИЯ_ФАКТИЧЕСКАЯ = "Сержанты"
    ElseIf (Значение = "ст. пр-к") _
        Or (Значение = "пр-к") Then
    КАТЕГОРИЯ_ФАКТИЧЕСКАЯ = "Прапорщики"
    ElseIf (Значение = "ряд.") _
        Or (Значение = "ефр.") Then
    КАТЕГОРИЯ_ФАКТИЧЕСКАЯ = "Солдаты"
    End If
    
    '    If (Должность = "Старшина") Or _
'        (Должность = "Старший техник") Then
'                КАТЕГОРИЯ_ФАКТИЧЕСКАЯ = "Прапорщики"
End Function

'Function ЦВЕТЯЧЕЙКИ(Ячейка As Range) As Variant
'    Dim CurrColor
'    CurrColor = Ячейка.Cells.Interior.Color
'    ЦВЕТЯЧЕЙКИ = CurrColor
'End Function

Function СЧЕТСТРОКПОУСЛ(Область_данных As Range, Тип_Военнослужащего As String) As Variant
    Dim Sum As Long
    Dim MaxCells As Long
    
    MaxCells = Application.WorksheetFunction.Max( _
        Application.Caller.Cells.Count, Область_данных.Cells.Count)
    ReDim Result(1 To MaxCells, 1 To 1)
    For Each Rng In Область_данных.Cells
        If Rng.Value = Тип_Военнослужащего Then
            Sum = Sum + 1
        End If
    Next Rng
    
    СЧЕТСТРОКПОУСЛ = Sum
    
End Function

Function РазвернутаяСтроеваяОтрывПрочие(Область_данных_подразделение As Range, Область_данных_отрыв As Range, ПОДРАЗДЕЛЕНИЕ As String) As Variant
    Dim Sum As Long
    Dim MaxCells As Long
    
    ' MaxCells = Application.WorksheetFunction.Max( _
    '     Application.Caller.Cells.Count, Область_данных_подразделение.Cells.Count)
    ' ReDim Result(1 To MaxCells, 1 To 1)
'    For Each Rng In Область_данных_подразделение.Cells
'        If Rng.Value = Подразделение Then
'             Sum = Sum + 1
'            For Each Rng2 In Область_данных_отрыв.Cells
'                If Rng2.Value = "н/о" Then
'                    Sum = Sum + 1
'                End If
'                Next Rng2
'        End If
'    Next Rng
    
    For Each Rng In Область_данных_отрыв.Cells
        If Rng.Value = "н/о" Then
            For Each Rng2 In Область_данных_подразделение.Cells
                If Rng2.Value = ПОДРАЗДЕЛЕНИЕ Then
                    Sum = Sum + 1
                End If
                Next Rng2
        End If
    Next Rng
    
    РазвернутаяСтроеваяОтрывПрочие = Sum
    
End Function

Function УБРПУСТЫЕ(Область_данных As Range) As Variant
    Dim N As Long
    Dim N2 As Long
    Dim Rng As Range
    Dim MaxCells As Long
    Dim Result() As Variant
    Dim R As Long
    Dim C As Long
    
    MaxCells = Application.WorksheetFunction.Max( _
        Application.Caller.Cells.Count, Область_данных.Cells.Count)
    ReDim Result(1 To MaxCells, 1 To 1)
    For Each Rng In Область_данных.Cells
        If Rng.Value <> vbNullString Then
            N = N + 1
            Result(N, 1) = Rng.Value
        End If
    Next Rng
    For N2 = N + 1 To MaxCells
        Result(N2, 1) = vbNullString
    Next N2
    
    If Application.Caller.Rows.Count = 1 Then
        УБРПУСТЫЕ = Application.WorksheetFunction.Transpose(Result)
    Else
        УБРПУСТЫЕ = Result
    End If
End Function

Public Function СЛДЕНЬ(День As Integer, НачалоНедели As Integer) As String
    Select Case НачалоНедели
    Case 1
        Select Case День
        Case 1
            СЛДЕНЬ = "Понедельник"
        Case 2
            СЛДЕНЬ = "Вторник"
        Case 3
            СЛДЕНЬ = "Среда"
        Case 4
            СЛДЕНЬ = "Четверг"
        Case 5
            СЛДЕНЬ = "Пятница"
        Case 6
            СЛДЕНЬ = "Суббота"
        Case 7
            СЛДЕНЬ = "Воскресенье"
        Case Else
            СЛДЕНЬ = "Некорректное значение для дня недели [1..7]"
        End Select
            Case Else
        СЛДЕНЬ = "Некорректное значение для Начала недели [1..12]"
    End Select
End Function
        
Public Function СЛМЕСЯЦ(Месяц As Integer, Падеж As Integer) As String

    Dim arrMonth(1 To 12, 1 To 6) As Variant
    '
    ' Заполнение массива
    ' вида (х,у), где
    '           х - месяц
    '           y - падеж
    '
    arrMonth(1, 1) = "Январь"
    arrMonth(1, 2) = "Января"
    arrMonth(1, 3) = "Январю"
    arrMonth(1, 4) = "Январь"
    arrMonth(1, 5) = "Январем"
    arrMonth(1, 6) = "Январе"
    
    arrMonth(2, 1) = "Февраль"
    arrMonth(2, 2) = "Февраля"
    arrMonth(2, 3) = "Февралю"
    arrMonth(2, 4) = "Февраль"
    arrMonth(2, 5) = "Февралем"
    arrMonth(2, 6) = "Феврале"
    
    arrMonth(3, 1) = "Март"
    arrMonth(3, 2) = "Марта"
    arrMonth(3, 3) = "Марту"
    arrMonth(3, 4) = "Март"
    arrMonth(3, 5) = "Мартом"
    arrMonth(3, 6) = "Марте"
    
    arrMonth(4, 1) = "Апрель"
    arrMonth(4, 2) = "Апреля"
    arrMonth(4, 3) = "Апрелю"
    arrMonth(4, 4) = "Апрель"
    arrMonth(4, 5) = "Апрелем"
    arrMonth(4, 6) = "Апреле"
    
    arrMonth(5, 1) = "Май"
    arrMonth(5, 2) = "Мая"
    arrMonth(5, 3) = "Маю"
    arrMonth(5, 4) = "Май"
    arrMonth(5, 5) = "Маем"
    arrMonth(5, 6) = "Мае"
    
    arrMonth(6, 1) = "Июнь"
    arrMonth(6, 2) = "Июня"
    arrMonth(6, 3) = "Июню"
    arrMonth(6, 4) = "Июнь"
    arrMonth(6, 5) = "Июнем"
    arrMonth(6, 6) = "Июне"
    
    arrMonth(7, 1) = "Июль"
    arrMonth(7, 2) = "Июля"
    arrMonth(7, 3) = "Июлю"
    arrMonth(7, 4) = "Июль"
    arrMonth(7, 5) = "Июлем"
    arrMonth(7, 6) = "Июле"
    
    arrMonth(8, 1) = "Август"
    arrMonth(8, 2) = "Августа"
    arrMonth(8, 3) = "Августу"
    arrMonth(8, 4) = "Август"
    arrMonth(8, 5) = "Августом"
    arrMonth(8, 6) = "Августе"
    
    arrMonth(9, 1) = "Сентябрь"
    arrMonth(9, 2) = "Сентября"
    arrMonth(9, 3) = "Сентябрю"
    arrMonth(9, 4) = "Сентябрь"
    arrMonth(9, 5) = "Сентябрем"
    arrMonth(9, 6) = "Сентябре"
    
    arrMonth(10, 1) = "Октябрь"
    arrMonth(10, 2) = "Октября"
    arrMonth(10, 3) = "Октябрю"
    arrMonth(10, 4) = "Октябрь"
    arrMonth(10, 5) = "Октябрем"
    arrMonth(10, 6) = "Октябре"
    
    arrMonth(11, 1) = "Ноябрь"
    arrMonth(11, 2) = "Ноября"
    arrMonth(11, 3) = "Ноябрю"
    arrMonth(11, 4) = "Ноябрь"
    arrMonth(11, 5) = "Ноябрем"
    arrMonth(11, 6) = "Ноябре"
   
    arrMonth(12, 1) = "Декабрь"
    arrMonth(12, 2) = "Декабря"
    arrMonth(12, 3) = "Декабрю"
    arrMonth(12, 4) = "Декабрь"
    arrMonth(12, 5) = "Декабрем"
    arrMonth(12, 6) = "Декабре"
    
    СЛМЕСЯЦ = arrMonth(Месяц, Падеж)
    
End Function

Function ПОЛНАЯ_ДАТА(КОГДА As Variant, ВИД As Integer) As String
    
    Dim sDay As String
    Dim sMonth As String
    Dim sYear As String
        
    КОГДА = StrConv(Trim(КОГДА), vbUpperCase)
    
    Select Case КОГДА
        Case "СЕГОДНЯ"
            dDate = 0
        Case "ЗАВТРА"
            dDate = 1
        Case Else:
            If Conversion.Int(КОГДА) >= 0 Then
                dDate = Conversion.Int(КОГДА)
            End If
    End Select
    
    sDay = CStr(DatePart("d", DateAdd("d", dDate, Now), vbMonday))
    sMonth = StrConv(CStr(СЛМЕСЯЦ(DatePart("m", DateAdd("d", dDate, Now), vbMonday), 2)), vbLowerCase)
    sYear = CStr(DatePart("yyyy", DateAdd("d", dDate, Now), vbMonday))
    
    If (КОГДА = "СЕГОДНЯ") Then
        Select Case ВИД
            Case 1
                ПОЛНАЯ_ДАТА = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "года"
            Case 2
                ПОЛНАЯ_ДАТА = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "г."
            Case 3
                ПОЛНАЯ_ДАТА = "«" + sDay + "»" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "года"
            Case 4
                ПОЛНАЯ_ДАТА = "«" + sDay + "»" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "г."
        End Select
        ElseIf (КОГДА = "ЗАВТРА") Then
            Select Case ВИД
                Case 1
                    ПОЛНАЯ_ДАТА = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "года"
                Case 2
                    ПОЛНАЯ_ДАТА = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "г."
                Case 3
                    ПОЛНАЯ_ДАТА = "«" + sDay + "»" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "года"
                Case 4
                    ПОЛНАЯ_ДАТА = "«" + sDay + "»" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "г."
            End Select
            Else:
                Select Case ВИД
                    Case 1
                        ПОЛНАЯ_ДАТА = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "года"
                    Case 2
                        ПОЛНАЯ_ДАТА = sDay + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "г."
                    Case 3
                        ПОЛНАЯ_ДАТА = "«" + sDay + "»" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "года"
                    Case 4
                        ПОЛНАЯ_ДАТА = "«" + sDay + "»" + Chr(32) + sMonth + Chr(32) + sYear + Chr(32) + "г."
                End Select
    End If
    
End Function

Public Function ВЫСЛУГА(Начальная_Дата As Date, Конечная_Дата As Date) As String
    Dim хГод As Integer
    Dim хМесяц  As Integer
    Dim хДень  As Integer
    
    Dim словоГод As String
    Dim СЛМЕСЯЦ As String
    Dim словоДень As String
    
    хГод = WorksheetFunction.RoundDown(((Конечная_Дата - Начальная_Дата) / 365), 0)
    хМесяц = WorksheetFunction.RoundDown(((((Конечная_Дата - Начальная_Дата) / 365) - хГод) * 12), 0)
    хДень = WorksheetFunction.RoundDown(((((((Конечная_Дата - Начальная_Дата) / 365) - хГод) * 12) - хМесяц) * (365 / 12)), 0)
    
    If (хГод = 1) Then
        словоГод = " год"
        ElseIf (хГод >= 2) And (хГод < 5) Then
        словоГод = " года"
            ElseIf (хГод >= 5) Or (хГод = 0) Then
            словоГод = " лет"
    End If
    
    If (хМесяц = 1) Then
        СЛМЕСЯЦ = " месяц"
        ElseIf (хМесяц >= 2) And (хМесяц < 5) Then
        СЛМЕСЯЦ = " месяца"
            ElseIf (хМесяц >= 5) Or (хМесяц = 0) Then
            СЛМЕСЯЦ = " месяцев"
    End If
    
    If (хДень = 1) Then
        словоДень = " день"
        ElseIf (хДень >= 2) And (хДень < 5) Then
        словоДень = " дня"
            ElseIf (хДень >= 5) Or (хДень = 0) Then
            словоДень = " дней"
    End If

    If хГод = 0 Then
        ВЫСЛУГА = CStr(хМесяц) + СЛМЕСЯЦ + " " + CStr(хДень) + словоДень
        ElseIf хМесяц = 0 And хДень = 0 Then
            ВЫСЛУГА = CStr(хГод) + словоГод
            ElseIf хМесяц = 0 Then
            ВЫСЛУГА = CStr(хГод) + словоГод + " " + CStr(хДень) + словоДень
                ElseIf хДень = 0 Then
                ВЫСЛУГА = CStr(хГод) + словоГод + " " + CStr(хМесяц) + СЛМЕСЯЦ
                Else
                ВЫСЛУГА = CStr(хГод) + словоГод + " " + CStr(хМесяц) + СЛМЕСЯЦ + " " + CStr(хДень) + словоДень
    End If
End Function

Public Function СУММВЫСЛУГА(Область_данных As Range) As Variant
    Dim хГод As Integer
    Dim хМесяц  As Integer
    Dim хДень  As Integer
    
    Dim Значение As Range
    ' ЧАСТНОЕ()
    ' WorksheetFunction.Quotient
    '
    ' ОСТАТ ()
    ' WorksheetFunction.Mod()
    
    For Each Значение In Область_данных.Cells
        хГод = хГод + WorksheetFunction.Quotient(хМесяц, 12)
        хМесяц = WorksheetFunction.Mod(хМесяц, 12) + WorksheetFunction.Quotient(хДень, 30)
        хДень = WorksheetFunction.Mod(хДень, 30)
    Next Значение
   
End Function

Public Function РазвернутаяСтроеваяПодразделение(ПОДРАЗДЕЛЕНИЕ As String) As String
    '
    ' 1 МСБ (г)
    '
    ПОДРАЗДЕЛЕНИЕ = LCase(ПОДРАЗДЕЛЕНИЕ)
    
    If ПОДРАЗДЕЛЕНИЕ = "управление 1мсб" Then
        РазвернутаяСтроеваяПодразделение = "Управление"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "1 мср" Then
        РазвернутаяСтроеваяПодразделение = "1 МСР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "2 мср" Then
        РазвернутаяСтроеваяПодразделение = "2 МСР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "3 мср" Then
        РазвернутаяСтроеваяПодразделение = "3 МСР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "минометная батарея" Then
        РазвернутаяСтроеваяПодразделение = "минбатр"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "1 огнеметный взвод" Then
        РазвернутаяСтроеваяПодразделение = "ОВ"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "1 развед взвод" Then
        РазвернутаяСтроеваяПодразделение = "РВ"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "1 взвод связи" Then
        РазвернутаяСтроеваяПодразделение = "ВС"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "1 взв.обеспеч." Then
        РазвернутаяСтроеваяПодразделение = "ВОб"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "1 мп" Then
        РазвернутаяСтроеваяПодразделение = "медпункт"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "управление 2мсб" Then
        РазвернутаяСтроеваяПодразделение = "Управление"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "4 мср" Then
        РазвернутаяСтроеваяПодразделение = "4 МСР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "5 мср" Then
        РазвернутаяСтроеваяПодразделение = "5 МСР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "6 мср" Then
        РазвернутаяСтроеваяПодразделение = "6 МСР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "2 минометная батарея" Then
        РазвернутаяСтроеваяПодразделение = "минбатр"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "2 огнеметный взвод" Then
        РазвернутаяСтроеваяПодразделение = "ОВ"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "2 развед взвод" Then
        РазвернутаяСтроеваяПодразделение = "РВ"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "2 взвод связи" Then
        РазвернутаяСтроеваяПодразделение = "ВС"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "2 взв.обеспеч." Then
        РазвернутаяСтроеваяПодразделение = "ВОб"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "2 мп" Then
        РазвернутаяСтроеваяПодразделение = "медпункт"
    '
    ' В распоряжении
    '
    ElseIf (ПОДРАЗДЕЛЕНИЕ = "распор.оф.") Or _
            (ПОДРАЗДЕЛЕНИЕ = "распор.пр-к.") Or _
            (ПОДРАЗДЕЛЕНИЕ = "распор.пр-к") Or _
            (ПОДРАЗДЕЛЕНИЕ = "распор.серж.") Or _
            (ПОДРАЗДЕЛЕНИЕ = "распор.серж.") Or _
            (ПОДРАЗДЕЛЕНИЕ = "распор.солд.") Or _
            (ПОДРАЗДЕЛЕНИЕ = "распор.солд") Then
        РазвернутаяСтроеваяПодразделение = "расп"
        
    '
    ' ср(с)
    '
    ElseIf ПОДРАЗДЕЛЕНИЕ = "ср(с)" Then
        РазвернутаяСтроеваяПодразделение = "СР(С)"
        
    '
    ' ГСАДн
    '
    ElseIf ПОДРАЗДЕЛЕНИЕ = "управление гсадн" Then
        РазвернутаяСтроеваяПодразделение = "упр. гсадн"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "1 гсабатр" Then
        РазвернутаяСтроеваяПодразделение = "1 гсабатр"
    
    ElseIf ПОДРАЗДЕЛЕНИЕ = "2 гсабатр" Then
        РазвернутаяСтроеваяПодразделение = "2 гсабатр"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "3 гсабатр" Then
        РазвернутаяСтроеваяПодразделение = "3 гсабатр"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "вуд" Then
        РазвернутаяСтроеваяПодразделение = "вуд"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "вод" Then
        РазвернутаяСтроеваяПодразделение = "вод"
    
    '
    ' Подразделения усиления
    '
    ElseIf ПОДРАЗДЕЛЕНИЕ = "зрабатр" Then
        РазвернутаяСтроеваяПодразделение = "зрабатр"
    
    ElseIf ПОДРАЗДЕЛЕНИЕ = "1 развед рота" Then
        РазвернутаяСтроеваяПодразделение = "рр"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "3 развед рота (аспн)" Then
        РазвернутаяСтроеваяПодразделение = "АСпН"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "исв" Then
        РазвернутаяСтроеваяПодразделение = "исв"
    
    ElseIf ПОДРАЗДЕЛЕНИЕ = "сводная пгс" Then
        РазвернутаяСтроеваяПодразделение = "ПГС"
    
    ElseIf ПОДРАЗДЕЛЕНИЕ = "взвод рэб" Then
        РазвернутаяСтроеваяПодразделение = "РЭБ"
       
    ElseIf ПОДРАЗДЕЛЕНИЕ = "рем.р." Then
        РазвернутаяСтроеваяПодразделение = "Рем.Рота"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "ро" Then
        РазвернутаяСтроеваяПодразделение = "РО"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "рмто" Then
        РазвернутаяСтроеваяПодразделение = "РМТО"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "мед.взвод" Then
        РазвернутаяСтроеваяПодразделение = "Мед. группа"
    
    ElseIf ПОДРАЗДЕЛЕНИЕ = "мед. группа" Then
        РазвернутаяСтроеваяПодразделение = "Мед. группа"
        
    '
    ' Приданные подразделения
    '
    ElseIf ПОДРАЗДЕЛЕНИЕ = "тр" Then
        РазвернутаяСтроеваяПодразделение = "ТР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "реабатр" Then
        РазвернутаяСтроеваяПодразделение = "РЕАБАТР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "вптур" Then
        РазвернутаяСтроеваяПодразделение = "вПТУР"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "впзрк" Then
        РазвернутаяСтроеваяПодразделение = "вПЗРК"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "ву реадн" Then
        РазвернутаяСтроеваяПодразделение = "ВУ РЕАДН"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "во реадн" Then
        РазвернутаяСтроеваяПодразделение = "ВО РЕАДН"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "бла" Then
        РазвернутаяСтроеваяПодразделение = "БПЛА"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "расчет бла" Then
        РазвернутаяСтроеваяПодразделение = "БПЛА"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "пан" Then
        РазвернутаяСтроеваяПодразделение = "ПАН"
        
    ElseIf ПОДРАЗДЕЛЕНИЕ = "ввп" Then
        РазвернутаяСтроеваяПодразделение = "ввп"
        
    '
    ' Для подразделения не определенного (или пустого)
    '
    Else: РазвернутаяСтроеваяПодразделение = "ОШИБКА"
    End If
End Function

Private Function СКЛОНЕНИЕ_ИМЯ(ДляСклонения As String, Падеж As Integer) As String
    Окончание = LCase(Mid(CStr(ДляСклонения), (Len(CStr(ДляСклонения)) - 1), Len(CStr(ДляСклонения))))
    Окончание_Расширенное = LCase(Mid(CStr(ДляСклонения), (Len(CStr(ДляСклонения)) - 2), Len(CStr(ДляСклонения))))
    Select Case Падеж
        ' Именительный
        Case 1
            СКЛОНЕНИЕ_ИМЯ = ДляСклонения
            
        ' Родительный
        Case 2
            ' Дмитрий, Георгий
            If (Окончание = "ий") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ия"
                    
            ' Павел
            ElseIf (Окончание = "ел") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ла"
                    
            ' Никита
            ElseIf (Окончание = "та") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ту"
                    
            ' Игорь, Сергей
            ElseIf (Окончание = "рь") Or _
                    (Окончание = "ль") Or _
                    (Окончание = "ай") Or _
                    (Окончание = "ей") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "я"
                    
            ' Илья
            ElseIf (Окончание = "ья") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "ю"
            
            ' Абдула
            ElseIf (Окончание = "ла") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "ы"
                
            ' Ханпаша
            ElseIf (Окончание = "ша") Or _
                    (Окончание = "за") Or _
                    (Окончание = "са") Or _
                    (Окончание = "фа") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "у"
                    
            '
            ElseIf (Окончание = "ан") Or _
                    (Окончание = "ат") Or _
                    (Окончание = "он") Or _
                    (Окончание = "ен") Or _
                    (Окончание = "ил") Or _
                    (Окончание = "им") Or _
                    (Окончание = "ег") Or _
                    (Окончание = "ем") Or _
                    (Окончание = "ав") Or _
                    (Окончание = "ик") Or _
                    (Окончание = "рт") Or _
                    (Окончание = "ам") Or _
                    (Окончание = "ын") Or _
                    (Окончание = "ер") Or _
                    (Окончание = "ор") Or _
                    (Окончание = "ид") Or _
                    (Окончание = "ал") Or _
                    (Окончание = "ек") Or _
                    (Окончание = "нд") Or _
                    (Окончание = "рд") Or _
                    (Окончание = "ед") Or _
                    (Окончание = "ир") Or _
                    (Окончание = "ар") Or _
                    (Окончание = "ин") Or _
                    (Окончание = "др") Then
                СКЛОНЕНИЕ_ИМЯ = ДляСклонения + "а"
             ElseIf (Окончание = "ис") Or _
                    (Окончание = "ур") Or _
                    (Окончание = "тр") Or _
                    (Окончание = "ад") Or _
                    (Окончание = "аз") Or _
                    (Окончание = "яс") Or _
                    (Окончание = "аш") Or _
                    (Окончание = "ул") Or _
                    (Окончание = "ош") Or _
                    (Окончание = "вр") Or _
                    (Окончание = "уф") Or _
                    (Окончание = "ус") Or _
                    (Окончание = "ет") Or _
                    (Окончание = "ес") Or _
                    (Окончание = "лл") Or _
                    (Окончание = "ас") Or _
                    (Окончание = "еб") Or _
                    (Окончание = "яр") Or _
                    (Окончание = "из") Or _
                    (Окончание = "оп") Or _
                    (Окончание = "ит") Or _
                    (Окончание = "рг") Or _
                    (Окончание = "ёд") Or _
                    (Окончание = "аб") Then
                СКЛОНЕНИЕ_ИМЯ = ДляСклонения + "а"
             ElseIf (Окончание = "ык") Or _
                    (Окончание = "кс") Or _
                    (Окончание = "ах") Or _
                    (Окончание = "ум") Or _
                    (Окончание = "иб") Then
                СКЛОНЕНИЕ_ИМЯ = ДляСклонения + "а"
            ElseIf (Окончание = "ли") Or _
                    (Окончание = "би") Or _
                    (Окончание = "зи") Or _
                    (Окончание = "лу") Or _
                    (Окончание = "ди") Or _
                    (Окончание = "жи") Then
                СКЛОНЕНИЕ_ИМЯ = ДляСклонения
            Else: СКЛОНЕНИЕ_ИМЯ = "!!!"
            End If
            
        ' Дательный
        Case 3
            ' Дмитрий, Георгий
            If (Окончание = "ий") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ию"
                    
            ' Павел
            ElseIf (Окончание = "ел") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "лу"
                    
            ' Никита
            ElseIf (Окончание = "та") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "те"
                    
            ' Игорь, Сергей
            ElseIf (Окончание = "рь") Or _
                    (Окончание = "ль") Or _
                    (Окончание = "ай") Or _
                    (Окончание = "ей") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "ю"
                    
            ' Илья
            ElseIf (Окончание = "ья") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "е"
            
            ' Абдула
            ElseIf (Окончание = "ла") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "е"
                
            ' Ханпаша
            ElseIf (Окончание = "ша") Or _
                    (Окончание = "за") Or _
                    (Окончание = "са") Or _
                    (Окончание = "фа") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "е"
                
            ElseIf (Окончание = "ан") Or _
                    (Окончание = "ат") Or _
                    (Окончание = "он") Or _
                    (Окончание = "ен") Or _
                    (Окончание = "им") Or _
                    (Окончание = "ег") Or _
                    (Окончание = "ем") Or _
                    (Окончание = "ав") Or _
                    (Окончание = "ик") Or _
                    (Окончание = "рт") Or _
                    (Окончание = "ам") Or _
                    (Окончание = "ын") Or _
                    (Окончание = "ер") Or _
                    (Окончание = "ор") Or _
                    (Окончание = "ид") Or _
                    (Окончание = "ал") Or _
                    (Окончание = "ек") Or _
                    (Окончание = "нд") Or _
                    (Окончание = "рд") Or _
                    (Окончание = "ед") Or _
                    (Окончание = "ир") Or _
                    (Окончание = "ар") Or _
                    (Окончание = "ин") Or _
                    (Окончание = "др") Then
                СКЛОНЕНИЕ_ИМЯ = ДляСклонения + "у"
             
             '
             ' Расширенные окончания 3-х буквенные
             '
             ' Даниле
             ElseIf (Окончание_Расширенное = "нил") Then
                    СКЛОНЕНИЕ_ИМЯ = ДляСклонения + "е"
            ' Любовь
             ElseIf (Окончание_Расширенное = "овь") Then
                    СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 3)) + "ви"
                    
             ' Карча
             ElseIf (Окончание_Расширенное = "рча") Or _
                    (Окончание_Расширенное = "мма") Or _
                    (Окончание_Расширенное = "жда") Or _
                    (Окончание_Расширенное = "нда") Or _
                    (Окончание_Расширенное = "джа") Or _
                    (Окончание_Расширенное = "лло") Or _
                    (Окончание_Расширенное = "лва") Then
                    СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "е"
                    
             ' Михаилу, Даниилу
             ElseIf (Окончание_Расширенное = "аил") Or _
                    (Окончание_Расширенное = "лум") Or _
                    (Окончание_Расширенное = "ход") Or _
                    (Окончание_Расширенное = "суп") Or _
                    (Окончание_Расширенное = "ияз") Or _
                    (Окончание_Расширенное = "зыр") Or _
                    (Окончание_Расширенное = "гос") Or _
                    (Окончание_Расширенное = "ков") Or _
                    (Окончание_Расширенное = "льф") Or _
                    (Окончание_Расширенное = "льд") Or _
                    (Окончание_Расширенное = "ият") Or _
                    (Окончание_Расширенное = "ейн") Or _
                    (Окончание_Расширенное = "сеп") Or _
                    (Окончание_Расширенное = "муд") Or _
                    (Окончание_Расширенное = "иил") Then
                    СКЛОНЕНИЕ_ИМЯ = ДляСклонения + "у"
             
             ' Чепелеу
             ElseIf (Окончание_Расширенное = "леу") Or _
                    (Окончание_Расширенное = "сто") Or _
                    (Окончание_Расширенное = "ами") Or _
                    (Окончание_Расширенное = "инэ") Or _
                    (Окончание_Расширенное = "ахи") Or _
                    (Окончание_Расширенное = "зри") Or _
                    (Окончание_Расширенное = "гжы") Or _
                    (Окончание_Расширенное = "мау") Then
                    СКЛОНЕНИЕ_ИМЯ = ДляСклонения
             
             ElseIf (Окончание = "ис") Or _
                    (Окончание = "ур") Or _
                    (Окончание = "тр") Or _
                    (Окончание = "ад") Or _
                    (Окончание = "аз") Or _
                    (Окончание = "яс") Or _
                    (Окончание = "аш") Or _
                    (Окончание = "ул") Or _
                    (Окончание = "ош") Or _
                    (Окончание = "вр") Or _
                    (Окончание = "уф") Or _
                    (Окончание = "ус") Or _
                    (Окончание = "ет") Or _
                    (Окончание = "ес") Or _
                    (Окончание = "лл") Or _
                    (Окончание = "ас") Or _
                    (Окончание = "еб") Or _
                    (Окончание = "яр") Or _
                    (Окончание = "из") Or _
                    (Окончание = "оп") Or _
                    (Окончание = "ит") Or _
                    (Окончание = "рг") Or _
                    (Окончание = "ёд") Or _
                    (Окончание = "аб") Then
                СКЛОНЕНИЕ_ИМЯ = ДляСклонения + "у"
             ElseIf (Окончание = "ык") Or _
                    (Окончание = "кс") Or _
                    (Окончание = "ах") Or _
                    (Окончание = "аф") Or _
                    (Окончание = "иб") Then
                СКЛОНЕНИЕ_ИМЯ = ДляСклонения + "у"
            ElseIf (Окончание = "ли") Or _
                    (Окончание = "би") Or _
                    (Окончание = "зи") Or _
                    (Окончание = "лу") Or _
                    (Окончание = "ди") Or _
                    (Окончание = "жи") Then
                СКЛОНЕНИЕ_ИМЯ = ДляСклонения
            
            ' Женские имена
            ElseIf (Окончание = "на") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "не"            '
            ElseIf (Окончание = "ся") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "се"
            ElseIf (Окончание = "ия") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ие"
            ElseIf (Окончание = "га") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ге"
            ElseIf (Окончание = "ра") Then
                СКЛОНЕНИЕ_ИМЯ = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ре"
                
            Else: СКЛОНЕНИЕ_ИМЯ = "!!!"
            End If
                       
        ' Винительный
        Case 4
            'СКЛОНЕНИЕ_ИМЯ =
            
        ' Творительный
        Case 5
            'СКЛОНЕНИЕ_ИМЯ =
        
        ' Предложный
        Case 6
            'СКЛОНЕНИЕ_ИМЯ =
        Case Else
            СКЛОНЕНИЕ_ИМЯ = "Некорректное значение для падежа [1..6]"
    End Select
End Function

Private Function СКЛОНЕНИЕ_ОТЧЕСТВО(ДляСклонения As String, Падеж As Integer) As String
    Окончание = LCase(Mid(CStr(ДляСклонения), (Len(CStr(ДляСклонения)) - 2), Len(CStr(ДляСклонения))))
    Select Case Падеж
        ' Именительный
        Case 1
            СКЛОНЕНИЕ_ОТЧЕСТВО = ДляСклонения
            
        ' Родительный
        Case 2
            If (Окончание = "вич") Then
                СКЛОНЕНИЕ_ОТЧЕСТВО = ДляСклонения + "а"
            ElseIf (Окончание = "ови") Or _
                    (Окончание = "ови") Or _
                    (Окончание = "аир") Or _
                    (Окончание = "ерт") Or _
                    (Окончание = "аги") Then
                СКЛОНЕНИЕ_ОТЧЕСТВО = ДляСклонения
            End If
        ' Дательный
        Case 3
            If (Окончание = "вич") Then
                СКЛОНЕНИЕ_ОТЧЕСТВО = ДляСклонения + "у"
            ElseIf (Окончание = "ови") Or _
                    (Окончание = "ови") Or _
                    (Окончание = "аир") Or _
                    (Окончание = "ерт") Or _
                    (Окончание = "глы") Or _
                    (Окончание = "ызы") Or _
                    (Окончание = "аги") Then
                СКЛОНЕНИЕ_ОТЧЕСТВО = ДляСклонения
            ' Женские отчества
            ElseIf (Окончание = "вна") Then
                СКЛОНЕНИЕ_ОТЧЕСТВО = (Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "не"
            End If
            
        ' Винительный
        Case 4
            'СКЛОНЕНИЕ_ОТЧЕСТВО =
            
        '
        Case 5
            'СКЛОНЕНИЕ_ОТЧЕСТВО =
            
        '
        Case 6
            'СКЛОНЕНИЕ_ОТЧЕСТВО =
            
        '
        Case Else
            СКЛОНЕНИЕ_ОТЧЕСТВО = "Некорректное значение для падежа [1..6]"
    End Select
End Function

Private Function СКЛОНЕНИЕ_ФАМИЛИЯ(ДляСклонения As String, Падеж As Integer) As String
    If ((Len(CStr(ДляСклонения)) - 1) > 0) Then
        Окончание = LCase(Mid(CStr(ДляСклонения), (Len(CStr(ДляСклонения)) - 1), Len(CStr(ДляСклонения))))
        Окончание_Расширенное = LCase(Mid(CStr(ДляСклонения), (Len(CStr(ДляСклонения)) - 2), Len(CStr(ДляСклонения))))
    Else:
        СКЛОНЕНИЕ_ФАМИЛИЯ = ДляСклонения
    End If
    
    Select Case Падеж
        ' Именительный
        Case 1
            СКЛОНЕНИЕ_ФАМИЛИЯ = ДляСклонения
            
        ' Родительный
        Case 2
            If (Окончание = "ов") Or _
                (Окончание = "ак") Or _
                (Окончание = "ин") Or _
                (Окончание = "ич") Or _
                (Окончание = "ос") Or _
                (Окончание = "ик") Or _
                (Окончание = "ук") Or _
                (Окончание = "ар") Or _
                (Окончание = "юк") Or _
                (Окончание = "ач") Or _
                (Окончание = "ян") Or _
                (Окончание = "ун") Or _
                (Окончание = "нц") Or _
                (Окончание = "ан") Or _
                (Окончание = "ол") Or _
                (Окончание = "як") Or _
                (Окончание = "ок") Or _
                (Окончание = "ёв") Or _
                (Окончание = "ык") Or _
                (Окончание = "ев") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения + "а")
            ElseIf (Окончание = "рь") Or _
                (Окончание = "ой") Or _
                (Окончание = "ый") Or _
                (Окончание = "ий") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ого")
            ElseIf (Окончание = "рь") Or _
                (Окончание = "ль") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "я")
            ElseIf (Окончание = "ца") Or _
                (Окончание = "га") Or _
                (Окончание = "ка") Or _
                (Окончание = "да") Or _
                (Окончание = "та") Or _
                (Окончание = "ва") Or _
                (Окончание = "да") Or _
                (Окончание = "за") Or _
                (Окончание = "на") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "у")
            ElseIf (Окончание = "ко") Or _
                (Окончание = "до") Or _
                (Окончание = "бо") Or _
                (Окончание = "ых") Or _
                (Окончание = "их") Or _
                (Окончание = "ло") Or _
                (Окончание = "ба") Or _
                (Окончание = "уш") Or _
                (Окончание = "са") Or _
                (Окончание = "ли") Or _
                (Окончание = "бу") Or _
                (Окончание = "нь") Or _
                (Окончание = "он") Or _
                (Окончание = "ла") Or _
                (Окончание = "це") Or _
                (Окончание = "жи") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения)
                
            ' Расширенные окончания
            ElseIf (Окончание_Расширенное = "так") Or _
                   (Окончание_Расширенное = "сак") Or _
                   (Окончание_Расширенное = "цзю") Or _
                   (Окончание_Расширенное = "мша") Or _
                   (Окончание_Расширенное = "цао") Or _
                   (Окончание_Расширенное = "кер") Or _
                   (Окончание_Расширенное = "иец") Or _
                   (Окончание_Расширенное = "чия") Or _
                   (Окончание_Расширенное = "дзе") Or _
                   (Окончание_Расширенное = "она") Or _
                   (Окончание_Расширенное = "леб") Or _
                   (Окончание_Расширенное = "рей") Or _
                   (Окончание_Расширенное = "оша") Or _
                   (Окончание_Расширенное = "нго") Or _
                   (Окончание_Расширенное = "рах") Or _
                   (Окончание_Расширенное = "идт") Or _
                   (Окончание_Расширенное = "дыч") Or _
                   (Окончание_Расширенное = "аха") Or _
                   (Окончание_Расширенное = "оюн") Or _
                   (Окончание_Расширенное = "руг") Or _
                   (Окончание_Расширенное = "льц") Or _
                   (Окончание_Расширенное = "ури") Or _
                   (Окончание_Расширенное = "епа") Or _
                   (Окончание_Расширенное = "вей") Or _
                   (Окончание_Расширенное = "мак") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения)
            ' Дюбель, Филь
            ElseIf (Окончание_Расширенное = "ель") Or _
                   (Окончание_Расширенное = "иль") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения)
                
            Else: СКЛОНЕНИЕ_ФАМИЛИЯ = "!!!"
            End If
        ' Дательный
        Case 3
            If (Окончание = "ов") Or _
                (Окончание = "ин") Or _
                (Окончание = "ич") Or _
                (Окончание = "ос") Or _
                (Окончание = "ик") Or _
                (Окончание = "ук") Or _
                (Окончание = "ар") Or _
                (Окончание = "юк") Or _
                (Окончание = "ач") Or _
                (Окончание = "ян") Or _
                (Окончание = "ун") Or _
                (Окончание = "нц") Or _
                (Окончание = "ан") Or _
                (Окончание = "ол") Or _
                (Окончание = "як") Or _
                (Окончание = "ёв") Or _
                (Окончание = "ык") Or _
                (Окончание = "ев") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения + "у")
            ElseIf (Окончание = "ый") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ю")
               
            ElseIf (Окончание = "ой") Or _
                (Окончание = "ий") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ому")
                
            ElseIf (Окончание = "ца") Or _
                (Окончание = "га") Or _
                (Окончание = "ка") Or _
                (Окончание = "да") Or _
                (Окончание = "та") Or _
                (Окончание = "да") Or _
                (Окончание = "за") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)))
            ElseIf (Окончание = "ко") Or _
                (Окончание = "до") Or _
                (Окончание = "бо") Or _
                (Окончание = "ых") Or _
                (Окончание = "их") Or _
                (Окончание = "ло") Or _
                (Окончание = "ба") Or _
                (Окончание = "уш") Or _
                (Окончание = "са") Or _
                (Окончание = "ли") Or _
                (Окончание = "бу") Or _
                (Окончание = "нь") Or _
                (Окончание = "он") Or _
                (Окончание = "ла") Or _
                (Окончание = "це") Or _
                (Окончание = "ещ") Or _
                (Окончание = "от") Or _
                (Окончание = "но") Or _
                (Окончание = "жи") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения)
            
            ' Женские фамилии
            ElseIf (Окончание = "ва") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "ой")
                
            ' Расширенные окончания
            ElseIf (Окончание_Расширенное = "так") Or _
                   (Окончание_Расширенное = "сак") Or _
                   (Окончание_Расширенное = "цзю") Or _
                   (Окончание_Расширенное = "мша") Or _
                   (Окончание_Расширенное = "цао") Or _
                   (Окончание_Расширенное = "кер") Or _
                   (Окончание_Расширенное = "иец") Or _
                   (Окончание_Расширенное = "чия") Or _
                   (Окончание_Расширенное = "дзе") Or _
                   (Окончание_Расширенное = "она") Or _
                   (Окончание_Расширенное = "леб") Or _
                   (Окончание_Расширенное = "рей") Or _
                   (Окончание_Расширенное = "оша") Or _
                   (Окончание_Расширенное = "нго") Or _
                   (Окончание_Расширенное = "рах") Or _
                   (Окончание_Расширенное = "идт") Or _
                   (Окончание_Расширенное = "дыч") Or _
                   (Окончание_Расширенное = "аха") Or _
                   (Окончание_Расширенное = "оюн") Or _
                   (Окончание_Расширенное = "руг") Or _
                   (Окончание_Расширенное = "льц") Or _
                   (Окончание_Расширенное = "ури") Or _
                   (Окончание_Расширенное = "епа") Or _
                   (Окончание_Расширенное = "вей") Or _
                   (Окончание_Расширенное = "мак") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения)
            ' Дюбель, Филь
            ElseIf (Окончание_Расширенное = "ель") Or _
                   (Окончание_Расширенное = "иль") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения)
            
            ' Лисицын
            ElseIf (Окончание_Расширенное = "цын") Or _
                   (Окончание_Расширенное = "гер") Or _
                   (Окончание_Расширенное = "коз") Or _
                   (Окончание_Расширенное = "дод") Or _
                   (Окончание_Расширенное = "чак") Or _
                   (Окончание_Расширенное = "мас") Or _
                   (Окончание_Расширенное = "льф") Or _
                   (Окончание_Расширенное = "дус") Or _
                   (Окончание_Расширенное = "лер") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения + "у")
            
            ' Дундера
            ElseIf (Окончание_Расширенное = "дер") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase(ДляСклонения + "е")
            
            ' Музыра, Кравчина
            ElseIf (Окончание_Расширенное = "ыра") Or _
                   (Окончание_Расширенное = "ина") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "е")
            
            ' Логутенок
            ElseIf (Окончание_Расширенное = "нок") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ку")
            
            ' Водождок
            ElseIf (Окончание_Расширенное = "док") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "ку")
                
            ' Кобзарь
            ElseIf (Окончание_Расширенное = "арь") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "ю")
            
            ' Король
            ElseIf (Окончание_Расширенное = "оль") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 1)) + "ю")
                
            ' Женские фамилии
            ElseIf (Окончание_Расширенное = "кая") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ой")
            ElseIf (Окончание_Расширенное = "няя") Then
                СКЛОНЕНИЕ_ФАМИЛИЯ = UCase((Mid(ДляСклонения, 1, Len(ДляСклонения) - 2)) + "ей")
                
                Else: СКЛОНЕНИЕ_ФАМИЛИЯ = "!!!"
            End If
       
        ' Винительный
        Case 4
            СКЛОНЕНИЕ_ФАМИЛИЯ = Окончание
        
        ' Творительный
        Case 5
            'СКЛОНЕНИЕ_ФАМИЛИЯ =
        
        ' Предложный
        Case 6
            'СКЛОНЕНИЕ_ФАМИЛИЯ =
        Case Else
            СКЛОНЕНИЕ_ФАМИЛИЯ = "Некорректное значение для падежа [1..6]"
    End Select
End Function

Function УНИФИКАЦИЯ(Значение As String) As Variant
    
    Dim buffLastName As String
    Dim buffFirstName As String
    Dim buffMidName As String
    Dim buffAppendName As String
    
    Dim buffSplit() As String
    
    Dim t As Integer
 
    ' Убрать пробелы слева и справа
    Значение = Trim(Значение)
    ' Убрать двойные пробелы
    Do While InStr(1, Значение, Space(2), 1) <> 0
        Значение = Replace(Значение, Space(2), Space(1), vbTextCompare)
    Loop
    
    ' Разбить строку с разделителем " "
    buffSplit() = Split(Trim(Значение), Chr(32))
    
    ' Получить значение - Фамилия
    buffLastName = buffSplit(0)
    ' Получить значение - Имя
    buffFirstName = buffSplit(1)
    ' Получить значение - Отчество
    buffMidName = buffSplit(2)
    
    t = 1
    
    On Error GoTo C
        buffAppendName = buffSplit(3)
        t = 2
C:
    
    Select Case t
        Case 1
            УНИФИКАЦИЯ = StrConv(buffLastName, vbUpperCase) + " " + StrConv(buffFirstName, vbProperCase) + " " + StrConv(buffMidName, vbProperCase)
            t = 0
        Case 2
            УНИФИКАЦИЯ = StrConv(buffLastName, vbUpperCase) + " " + StrConv(buffFirstName, vbProperCase) + " " + StrConv(buffMidName, vbProperCase) + " " + StrConv(buffAppendName, vbProperCase)
            t = 0
    End Select
End Function

Function ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ(Звание As String) As String
    If Звание = "п-к" Then
        ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "полковник"
        ElseIf Звание = "п/п-к" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "подполковник"
        ElseIf Звание = "м-р" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "майор"
        ElseIf Звание = "к-н" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "капитан"
        ElseIf Звание = "ст. л-т" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "старший лейтенант"
        ElseIf Звание = "л-т" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "лейтенант"
        ElseIf Звание = "мл. л-т" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "младший лейтенант"
        ElseIf Звание = "ст. пр-к" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "старший прапорщик"
        ElseIf Звание = "пр-к" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "прапорщик"
        ElseIf Звание = "ст-на" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "старшина"
        ElseIf Звание = "ст. с-т" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "старший сержант"
        ElseIf Звание = "с-т" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "сержант"
        ElseIf Звание = "мл. с-т" Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "младший сержант"
        ElseIf Звание = "ефр." Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "ефрейтор"
        ElseIf Звание = "ряд." Then
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "рядовой"
        Else:
            ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ = "!!!"
    End If
End Function

Function ПОДРАЗДЕЛЕНИЕ(Рота As String, Взвод As String, Отделение As String) As String

    Dim Тип As Integer
    Тип = 0
    
    If (Рота = "1 мср") Or _
       (Рота = "2 мср") Or _
       (Рота = "3 мср") Or _
       (Рота = "Минометная Батарея") Then
            Тип = 1
        ElseIf (Рота = "1 Огнеметный взвод") Or _
            (Рота = "1 Развед взвод") Or _
            (Рота = "1 взвод связи") Or _
            (Рота = "1 взв.обеспеч.") Or _
            (Рота = "1 МП") Then
                Тип = 2
    End If
    
    Select Case Тип
        Case 0
            If Рота = "управление 1мсб" Then
                    ПОДРАЗДЕЛЕНИЕ = "упр. 1 мсб(г)"
                ElseIf (Рота = "управление 1мсб") And (Взвод = "Штаб") Then
                    ПОДРАЗДЕЛЕНИЕ = "упр. 1 мсб(г)"
            End If
        Case 1
            ПОДРАЗДЕЛЕНИЕ = Рота
        Case 2
            ПОДРАЗДЕЛЕНИЕ = Взвод
    End Select
End Function



