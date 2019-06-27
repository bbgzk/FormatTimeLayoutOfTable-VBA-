Sub 调整轨迹查询记录格式()
'
' 宏1 宏
'

'
'调整轨迹查询记录格式，调整成“日期开始时间-结束时间”格式并排序

'    Dim i
'    For i = 1 To [65536].End(3).Row
'        If InStr(Cells(i, 1), "用户名：喻杰") = 0 Then
'            MsgBox i
'        End If
'    Next

'删除最后一行
    Range("A957:F957").Select
    Selection.Delete Shift:=xlUp

'排序数据，自编号顺序，时长逆序
    Range("A3:F956").Select
    ActiveWorkbook.Worksheets("轨迹查看日志").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("轨迹查看日志").Sort.SortFields.Add Key:=Range("A4:A956") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    ActiveWorkbook.Worksheets("轨迹查看日志").Sort.SortFields.Add Key:=Range("D4:D956") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("轨迹查看日志").Sort
        .SetRange Range("A3:F956")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'   删除第一二行
    Rows("1:2").Select
    Range("A2").Activate
    Selection.Delete Shift:=xlUp

'   在C列前插入5列
    Columns("C:C").Select
    Range("C3").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'分列数据“开始日期”
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        "-", FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2)), _
        TrailingMinusNumbers:=True

'分列数据“开始时间”
    Columns("E:E").Select
    Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2)), _
        TrailingMinusNumbers:=True
    ActiveWindow.SmallScroll Down:=-12

'   在I列前插入5列
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'分列数据“结束时间”
    Columns("H:H").Select
    Selection.TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        ":", FieldInfo:=Array(Array(1, 1), Array(2, 2), Array(3, 2), Array(4, 2)), _
        TrailingMinusNumbers:=True

'组合开始时间和结束时间成“日期开始时间-结束时间”样式
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=RC[-10]&RC[-9]&RC[-8]&RC[-7]&RC[-6]&RC[-5]&""-""&RC[-3]&RC[-2]&RC[-1]"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L954")
    Range("L2:L954").Select

'分列操作时间
    Columns("P:P").Select
    Selection.TextToColumns Destination:=Range("Q1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:= _
        "-", FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 1)), _
        TrailingMinusNumbers:=True

'组合操作日期
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=RC[4]&RC[5]&RC[6]"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M954")
    Range("M2:M954").Select
    
'把L和M列公式算出来的数据变成字符串

    Columns("L:L").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Columns("M:M").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'删除多余列
    Range("B:K,O:T").Select
    Selection.Delete Shift:=xlToLeft
    Range("B1").Select

End Sub
