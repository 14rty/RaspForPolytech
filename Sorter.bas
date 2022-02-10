Attribute VB_Name = "Module2"
Sub FinSort()
Attribute FinSort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FinSort Макрос
'

'
    
    Range("A1:I300").Select
    Range("A300").Activate
    ActiveWorkbook.Worksheets("Дисциплины").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Дисциплины").Sort.SortFields.Add Key:=Range( _
        "E2:E300"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "Первый семестр,Второй семестр,Третий семестр,Четвертый семестр,Пятый семестр,Шестой семестр,Седьмой семестр,Восьмой семестр,Девятый семестр,Десятый семестр,Одиннадцатый семестр" _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Дисциплины").Sort
        .SetRange Range("A1:I300")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
