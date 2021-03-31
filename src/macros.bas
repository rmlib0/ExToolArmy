Attribute VB_Name = "Macros"
'
' Сокращенные звания в полные
'
Sub ВПолныеВоинскиеЗвания()
Attribute ВПолныеВоинскиеЗвания.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' Подполковник
    '
    Cells.Replace What:="п/п-к", Replacement:="подполковник", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Полковник
    '
    Cells.Replace What:="п-к", Replacement:="полковник", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Майор
    '
    Cells.Replace What:="м-р", Replacement:="майор", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Капитан
    '
    Cells.Replace What:="к-н", Replacement:="капитан", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Старший лейтенант
    '
    Cells.Replace What:="ст. л-т", Replacement:="старший лейтенант", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Лейтенант
    '
    Cells.Replace What:="л-т", Replacement:="лейтенант", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Старший прапорщик / прапорщик
    '
    Cells.Replace What:="ст. пр-к", Replacement:="старший прапорщик", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="пр-к", Replacement:="прапорщик", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Старшина
    '
    Cells.Replace What:="ст-на", Replacement:="старшина", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Старший сержант
    '
    Cells.Replace What:="ст. с-т", Replacement:="старший сержант", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Младший сержант
    '
    Cells.Replace What:="мл. с-т", Replacement:="младший сержант", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    '
    ' Сержант
    '
    Cells.Replace What:="с-т", Replacement:="сержант", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Рядовой
    '
    Cells.Replace What:="ряд.", Replacement:="рядовой", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Ефрейтор
    '
    Cells.Replace What:="ефр.", Replacement:="ефрейтор", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

'
' Единое оформление воинских званий
'
Sub ОформлениеВоинскиеЗвания()
Attribute ОформлениеВоинскиеЗвания.VB_ProcData.VB_Invoke_Func = "Z\n14"
    '
    ' Подполковник
    '
    Cells.Replace What:="п/п-к.", Replacement:="п/п-к", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Полковник
    '
    Cells.Replace What:="п-к.", Replacement:="п-к", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Майор
    '
    Cells.Replace What:="м-р.", Replacement:="м-р", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="м - р.", Replacement:="м-р", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="м- р.", Replacement:="м-р", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="м -р.", Replacement:="м-р", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="м - р", Replacement:="м-р", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="м- р", Replacement:="м-р", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="м -р", Replacement:="м-р", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Капитан
    '
    Cells.Replace What:="к-н.", Replacement:="к-н", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="к- н.", Replacement:="к-н", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="к -н.", Replacement:="к-н", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="к - н.", Replacement:="к-н", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="к- н", Replacement:="к-н", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="к -н", Replacement:="к-н", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="к - н", Replacement:="к-н", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Старший лейтенант
    '
    Cells.Replace What:="ст. л-т.", Replacement:="ст. л-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ст.л-т", Replacement:="ст. л-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ст.л-т.", Replacement:="ст. л-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ст.л - т.", Replacement:="ст. л-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ст.л -т.", Replacement:="ст. л-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ст.л- т.", Replacement:="ст. л-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Лейтенант
    '
    Cells.Replace What:="л-т.", Replacement:="л-т", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="л- т.", Replacement:="л-т", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="л -т.", Replacement:="л-т", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="л - т.", Replacement:="л-т", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="л- т", Replacement:="л-т", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="л -т", Replacement:="л-т", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="л - т", Replacement:="л-т", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Старший прапорщик / прапорщик
    '
    Cells.Replace What:="ст.пр-к", Replacement:="ст. пр-к", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="пр-к.", Replacement:="пр-к", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Старшина
    '
    Cells.Replace What:="ст-на.", Replacement:="ст-на", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Старший сержант
    '
    Cells.Replace What:="ст.с-т", Replacement:="ст. с-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ст.с-т.", Replacement:="ст. с-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ст с-т.", Replacement:="ст. с-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ст с-т", Replacement:="ст. с-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Младший сержант
    '
    Cells.Replace What:="мл.с-т", Replacement:="мл. с-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="мл.с-т.", Replacement:="мл. с-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="мл с-т.", Replacement:="мл. с-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="мл с-т", Replacement:="мл. с-т", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Сержант
    '
    Cells.Replace What:="с-т.", Replacement:="с-т", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Рядовой
    '
    Cells.Replace What:="ряд .", Replacement:="ряд.", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '
    ' Ефрейтор
    '
    Cells.Replace What:="ефр .", Replacement:="ефр.", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

