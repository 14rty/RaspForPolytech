Attribute VB_Name = "Module2"
Sub FinSort()
Attribute FinSort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FinSort ������
'

'
    
    Range("A1:I300").Select
    Range("A300").Activate
    ActiveWorkbook.Worksheets("����������").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("����������").Sort.SortFields.Add Key:=Range( _
        "E2:E300"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "������ �������,������ �������,������ �������,��������� �������,����� �������,������ �������,������� �������,������� �������,������� �������,������� �������,������������ �������" _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("����������").Sort
        .SetRange Range("A1:I300")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
