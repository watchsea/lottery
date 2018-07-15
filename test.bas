Attribute VB_Name = "test"
Sub 程序升级()

    Dim sheet1 As Worksheet
    Dim tmClBgCol As Long       '总排名交锋等级数据所在列

    Set sheet1 = ActiveWorkbook.Sheets("综合数据")
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    
    
    '处理【主队+客队盘形分析(相对数据)】
    tmClBgCol = dataColDict.Item("ANARATIO_1")

    If tmClBgCol = 0 Then
        i1 = dataColDict.Item("ANARATIO")
        '修改原来的名称代码
        i2 = i1 + 4
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "胜"
        sheet1.Cells(3, i2 + 1) = "平"
        sheet1.Cells(3, i2 + 2) = "负"
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "主队+客队盘形分析(相对数据)-总排名"
        sheet1.Cells(4, i2) = "ANARATIO_1"
        
        MsgBox ("主队+客队盘形分析程序更新完毕！")
    Else
        MsgBox ("主队+客队盘分析程序已更新！")
    End If
    
    
    '处理bet365盘口评测
    tmClBgCol = dataColDict.Item("PANB_1")

    If tmClBgCol = 0 Then
        i1 = dataColDict.Item("PANB")
        '修改原来的名称代码
        i2 = i1 + 4
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "贴水"
        sheet1.Cells(3, i2 + 1) = "盘口"
        sheet1.Cells(3, i2 + 2) = "贴水"
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "bet365最新盘口评测"
        sheet1.Cells(4, i2) = "PANB_3"
        
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "贴水"
        sheet1.Cells(3, i2 + 1) = "盘口"
        sheet1.Cells(3, i2 + 2) = "贴水"
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "bet365赛前8小时盘口评测"
        sheet1.Cells(4, i2) = "PANB_2"
        
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "贴水"
        sheet1.Cells(3, i2 + 1) = "盘口"
        sheet1.Cells(3, i2 + 2) = "贴水"
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "bet365初始盘口评测"
        sheet1.Cells(4, i2) = "PANB_1"
        
        'sheet1.Cells.Columns(i2).ShrinkToFit = True
        MsgBox ("Bet365盘口评测程序更新完毕！")
    Else
        MsgBox ("Bet365盘口评测程序已更新！")
    End If
    
    
    '处理澳门盘口评测
    tmClBgCol = dataColDict.Item("PANM_1")
    If tmClBgCol = 0 Then
        i1 = dataColDict.Item("PANM")
        '修改原来的名称代码
        i2 = i1 + 4
        
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "贴水"
        sheet1.Cells(3, i2 + 1) = "盘口"
        sheet1.Cells(3, i2 + 2) = "贴水"
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "澳门彩票最新盘口评测"
        sheet1.Cells(4, i2) = "PANM_3"
        
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "贴水"
        sheet1.Cells(3, i2 + 1) = "盘口"
        sheet1.Cells(3, i2 + 2) = "贴水"
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "澳门彩票赛前8小时盘口评测"
        sheet1.Cells(4, i2) = "PANM_2"
        
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells(3, i2) = "贴水"
        sheet1.Cells(3, i2 + 1) = "盘口"
        sheet1.Cells(3, i2 + 2) = "贴水"
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "澳门彩票初始盘口评测"
        sheet1.Cells(4, i2) = "PANM_1"
                
        MsgBox ("澳门彩票盘口程序更新完毕！")
    Else
        MsgBox ("澳门彩票盘口程序已更新！")
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
