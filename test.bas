Attribute VB_Name = "test"
Sub 程序升级()

    Dim sheet1 As Worksheet
    Dim tmClBgCol As Long       '总排名交锋等级数据所在列

    Set sheet1 = ActiveWorkbook.Sheets("综合数据")
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    tmClBgCol = dataColDict.Item("SCHEMA")

    If tmClBgCol = 0 Then
        i1 = dataColDict.Item("SCHEMA") + 6
        sheet1.Cells.Columns(i1).Insert
        sheet1.Cells.Columns(i1).Insert
        sheet1.Cells(3, i1) = "模式7"
        sheet1.Cells(3, i1 + 1) = "模式8"
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
