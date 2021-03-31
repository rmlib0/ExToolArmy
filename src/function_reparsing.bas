Attribute VB_Name = "function_reparsing"
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

'
' StrPrepare
'
' Функция подготовки строки
'
Function StrPrepare(m As String) As String

    ' Убрать пробелы слева и справа
    m = Trim(m)
    
    ' Убрать двойные пробелы
    Do While InStr(1, m, Space(2), 1) <> 0
        m = Replace(m, Space(2), Space(1), vbTextCompare)
    Loop
    
    StrPrepare = m
    
End Function

'
' DTW_Month
'
' Функция перевода месяца в слово
'
Public Function DTW_Month(m As Integer, pcase As Integer) As String

    Dim arrJoin(1 To 12) As Variant
    
    arrJoin(1) = Array("Январь", "Января", "Январю", "Январь", "Январем", "Январе")
    arrJoin(2) = Array("Февраль", "Февраля", "Февралю", "Февраль", "Февралем", "Феврале")
    arrJoin(3) = Array("Март", "Марта", "Марту", "Март", "Мартом", "Марте")
    arrJoin(4) = Array("Апрель", "Апреля", "Апрелю", "Апрель", "Апрелем", "Апреле")
    arrJoin(5) = Array("Май", "Мая", "Маю", "Май", "Маем", "Мае")
    arrJoin(6) = Array("Июнь", "Июня", "Июню", "Июнь", "Июнем", "Июне")
    arrJoin(7) = Array("Июль", "Июля", "Июлю", "Июль", "Июлем", "Июле")
    arrJoin(8) = Array("Август", "Августа", "Августу", "Август", "Августом", "Августе")
    arrJoin(9) = Array("Сентябрь", "Сентября", "Сентябрю", "Сентябрь", "Сентябрем", "Сентябре")
    arrJoin(10) = Array("Октябрь", "Октября", "Октябрю", "Октябрь", "Октябрем", "Октябре")
    arrJoin(11) = Array("Ноябрь", "Ноября", "Ноябрю", "Ноябрь", "Ноябрем", "Ноябре")
    arrJoin(12) = Array("Декабрь", "Декабря", "Декабрю", "Декабрь", "Декабрем", "Декабре")

    DTW_Month = arrJoin(m)(pcase - 1)
    
End Function

'
' FaceBit
'
' Функция определения человека в отрыве
'
Function FaceBit(m As String) As Boolean
    
    arrAbsence = Array("командировка", _
                    "отпуск", _
                    "госпиталь", _
                    "лазарет", _
                    "болезнь", _
                    "медрота", _
                    "наряд", _
                    "ппд", _
                    "н/о", _
                    "рампа", _
                    "увольняется", _
                    "уволен", _
                    "полигон", _
                    "не передан", _
                    "выезд", _
                    "арест", _
                    "уголовное дело", _
                    "перевод", _
                    "2 мсб", _
                    "выходной", _
                    "работы" _
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
' Функция подготовки данных для ежедневного рапорта
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
    
    ' Получить количество строк
    NumRows = t.Rows.Count
    
    ' Получить количество столбцов
    NumColumns = t.Columns.Count
    
    ' Установить размеры массивов
    ReDim arrBuff(0 To NumColumns - 1)
    ReDim arrJoin(0 To NumRows - 1)

    
    For i = 1 To NumRows
        If StrConv(CStr(Trim(t(i, 1).Value)), vbLowerCase) = m Then
            arrBuff(0) = t(i, 2).Value
            arrBuff(1) = t(i, 3).Value
            arrBuff(2) = t(i, 4).Value
            
            ' Проверка даты возвращения
            ' Если не пустая или не равна нулю, тогда объединить
            If arrBuff(2) <> vbEmpty Or Format(arrBuff(2), "yyyy") > 2000 Then
                arrJoin(j) = ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ(CStr(arrBuff(0))) + " " + ИНИЦИАЛЫ(CStr(arrBuff(1)), 1) + " по " + CStr(arrBuff(2))
                
                ' Если дата возвращения пустая или равна нулю
                Else:
                    arrJoin(j) = ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ(CStr(arrBuff(0))) + " " + ИНИЦИАЛЫ(CStr(arrBuff(1)), 1)
            End If
            
            j = j + 1
            
            Else:
                RprtPrep = ""
        End If
    Next i
    
    ' Установить размеры массива
    ReDim arrPeople(0 To j - 1)
    
    ' Заполнить массив
    For k = 0 To UBound(arrPeople)
        arrPeople(k) = CStr(arrPeople(k)) + CStr(arrJoin(k))
    Next k

    ' Вернуть значение функции, согласно форматированию
    RprtPrep = StrConv(CStr(m), vbLowerCase) + " " + _
                Chr(151) + " " + CStr(j) + " чел. (" + _
                Join(arrPeople, ", ") + ")"

End Function

'
' RprtUnion
'
' Функция сборки для ежедневного рапорта
'
Function RprtUnion(d As Variant, t As Range) As Variant

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    Dim arrBuff() As Variant

    Dim NumRows As Integer

    d = StrConv(CStr(Trim(d)), vbLowerCase)
    
    ' Получить количество строк
    NumRows = t.Rows.Count
    
    ' Установить размеры массива
    ReDim arrBuff(1 To NumRows)
    
    j = 1
    
    For i = 1 To NumRows
        If StrConv(CStr(Trim(t(i, 1).Value)), vbLowerCase) <> d Then
        
            ' Если строка не разделитель, то занести значение в массив и увеличить счетчик
            arrBuff(j) = t(i, 1).Value
            j = j + 1
            
        End If
    Next i
    
    ReDim Preserve arrBuff(1 To j)
    
    RprtUnion = Join(arrBuff, Chr(10))
    
End Function

Private Function fRole(m As String) As String
    
    m = StrPrepare(m)
    
    arrRole_1 = Array("Командир батальона", _
                        "Заместитель командира батальона по ВПР", _
                        "Заместитель командира батальона по вооружению", _
                        "Помощник командира батальона по горной подготовке", _
                        "Начальник штаба - заместитель командира батальона", _
                        "Заместитель начальника штаба", _
                        "Командир роты", _
                        "Заместитель командира роты по ВПР", _
                        "Командир взвода", _
                        "Командир батареи", _
                        "Командир взвода - старший офицер на батарее", _
                        "Начальник связи - командир взвода" _
                        )
    For i = 0 To UBound(arrRole_1)
        If m = arrRole_1(i) Then
            fRole = sRole_1
            Exit For
        End If
    Next i
    
End Function

Private Function РОЛЬ_В_КАРАУЛЕ(Должность As String) As String
    Dim Роль As Integer
    
    Должность = StrPrepare(Должность)
    
        If Должность = "Командир батальона" Or _
            Должность = "Заместитель командира батальона по ВПР" Or _
            Должность = "Заместитель командира батальона по вооружению" Or _
            Должность = "Помощник командира батальона по горной подготовке" Or _
            Должность = "Начальник штаба - заместитель командира батальона" Or _
            Должность = "Заместитель начальника штаба" Or _
            Должность = "Командир роты" Or _
            Должность = "Заместитель командира роты по ВПР" Or _
            Должность = "Командир взвода" Or _
            Должность = "Командир батареи" Or _
            Должность = "Командир взвода - старший офицер на батарее" Or _
            Должность = "Начальник связи - командир взвода" Then
                Роль = 1
            ElseIf Должность = "Заместитель командира взвода - командир отделения" Or _
                    Должность = "Старшина" Or _
                    Должность = "Старший техник" Or _
                    Должность = "Заместитель командира взвода - командир миномета" Or _
                    Должность = "Командир отделения" Or _
                    Должность = "Командир миномета" Or _
                    Должность = "Командир отделения (подвоза боеприпасов)" Then
                Роль = 2
            Else:
                Роль = 3
        End If
    Select Case Роль
        Case 1
            РОЛЬ_В_КАРАУЛЕ = "начальник караула"
        Case 2
            РОЛЬ_В_КАРАУЛЕ = "помощник начальника караула, разводящий"
        Case 3
            РОЛЬ_В_КАРАУЛЕ = "караульный"
    End Select
End Function

Function ДЛЯ_ПРИКАЗА_НА_ДОПУСК(Рота As String, Взвод As String, Отделение As String, Должность As String, Звание As String, ФИО As String) As String
    
    Должность = Trim(Должность)
    Do While InStr(1, Должность, Space(2), 1) <> 0
        Должность = Replace(Должность, Space(2), Space(1), vbTextCompare)
    Loop
    
    ДЛЯ_ПРИКАЗА_НА_ДОПУСК = ( _
                                Должность + Chr(32) + _
                                ПОЛНОЕ_ВОИНСКОЕ_ЗВАНИЕ(Звание) + Chr(32) + _
                                ИНИЦИАЛЫ(StrConv(УНИФИКАЦИЯ(ФИО), vbProperCase), 1) + Chr(32) + Chr(150) + Chr(32) + _
                                РОЛЬ_В_КАРАУЛЕ(Должность) + ";")
End Function

