Attribute VB_Name = "test"
Sub 程序升级()

    Dim sheet1 As Worksheet
    Dim tmClBgCol As Long       '总排名交锋等级数据所在列

    Set sheet1 = ActiveWorkbook.Sheets("综合数据")
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    tmClBgCol = dataColDict.Item("SCHEMA4")

    If tmClBgCol = 0 Then
        i1 = dataColDict.Item("SCHEMA")
        '修改原来的名称代码
        For i = 4 To 8
            sheet1.Cells(4, i1 + i - 1) = "SCHEMA" & i
        Next
        '处理模式八数据
        i2 = i1 + 8
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "模式8值"
        sheet1.Cells(3, i2 + 1) = "模式8比较"
        sheet1.Cells(3, i2 + 2) = "五八比较"
        sheet1.Cells.Columns(i2).ShrinkToFit = True
        
        '处理模式七数据
        i2 = i1 + 7
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "模式7值"
        sheet1.Cells(3, i2 + 1) = "模式7比较"
        sheet1.Cells(3, i2 + 2) = "四七比较"
        sheet1.Cells.Columns(i2).ShrinkToFit = True
        
        '处理模式五数据
        i2 = i1 + 5
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "模式5值"
        sheet1.Cells(3, i2 + 1) = "模式5比较"
        sheet1.Cells.Columns(i2).ShrinkToFit = True
        
        '处理模式四数据
        i2 = i1 + 4
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "模式4值"
        sheet1.Cells(3, i2 + 1) = "模式4比较"
        sheet1.Cells.Columns(i2).ShrinkToFit = True
        
        MsgBox ("程序更新完毕！")
    Else
        MsgBox ("程序已更新！")
    End If
End Sub



Sub 网站数据导入测试()
Dim begindate As Date
Dim enddate As Date

    begindate = DateAdd("d", -1, Date)
    enddate = DateAdd("d", 2, Date)
    
    'Call 澳客网必发盈亏(begindate, enddate)
    'Call 澳客网胜负指数(begindate, enddate)
    'Call 澳客网盘口评测(begindate, enddate)
    'Call 澳客网凯利指数(begindate, enddate)
    Call 澳客网数据载入
    MsgBox ("导入完毕！")
End Sub
