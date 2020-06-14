Attribute VB_Name = "test"
Sub 程序升级(ByRef control As Office.IRibbonControl)
    Call 程序升级20200315
End Sub
Sub 程序升级20200315()
    Dim sheet1 As Worksheet
    Dim tmClBgCol As Long
    Dim colDesp
    Dim colIndex, cols
    Dim insertColIdx
    Dim k1, i, j
    Dim baseCnt
    Dim colRange As String

    Call 配置数据载入(dataConfig, "Config")
    Set sheet1 = ActiveWorkbook.Sheets("综合数据")
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    
    i = dataColDict.Item("LOSE1")  '数据插在LOSE1之前

     
    For k1 = 2 To UBound(dataConfig)
        If Not (dataColDict.exists(dataConfig(k1, 2))) And dataConfig(k1, 15) = "Y" And dataConfig(k1, 16) = "20200315" Then
            baseCnt = CInt(dataConfig(k1, 9))
            If baseCnt = 4 Then
                colIndex = "胜,平,负,返还率"
            Else
                colIndex = "胜,平,负"
            End If
            If CBool(dataConfig(k1, 7)) Then
                baseCnt = baseCnt + 1
                colIndex = colIndex + ",标识"
            End If
            If dataConfig(k1, 8) <> "FALSE" Then
                baseCnt = baseCnt + 1
                colIndex = colIndex + ",比较"
            End If
            
            For j = 1 To baseCnt
                sheet1.Cells.Columns(i).Insert Shift:=xlToRight   ', CopyOrigin:=xlFormatFromLeftOrAbove
            Next
            'sheet1.Cells.Columns(colRange).Insert Shift:=xlToRight   ', CopyOrigin:=xlFormatFromLeftOrAbove
            sheet1.Range(Cells(2, i), Cells(2, i + baseCnt - 1)).Merge
            colDesp = dataConfig(k1, 1)
            sheet1.Cells(2, i) = colDesp
            cols = Split(colIndex, ",")
            sheet1.Range(Cells(2, i), Cells(3, i + baseCnt - 1)).Borders.LineStyle = 1
            For j = 0 To baseCnt - 1
                sheet1.Cells(3, i + j) = cols(j)
            Next
            sheet1.Cells(4, i) = dataConfig(k1, 2)
            i = i + baseCnt
            
        End If
    Next
    
    '必发指数
    If Not (dataColDict.exists("BFZS")) Then
    
        For j = 1 To 3
            sheet1.Cells.Columns(i).Insert Shift:=xlToRight   ', CopyOrigin:=xlFormatFromLeftOrAbove
        Next
        'sheet1.Cells.Columns(colRange).Insert Shift:=xlToRight   ', CopyOrigin:=xlFormatFromLeftOrAbove
        sheet1.Range(Cells(2, i), Cells(2, i + 2)).Merge
        sheet1.Cells(2, i) = "必发指数"
        colIndex = "胜,平,负"
        cols = Split(colIndex, ",")
        sheet1.Range(Cells(2, i), Cells(3, i + baseCnt - 1)).Borders.LineStyle = 1
        For j = 0 To 2
            sheet1.Cells(3, i + j) = cols(j)
        Next
        sheet1.Cells(4, i) = "BFZS"
    End If
    
    MsgBox ("升级完成！")
    
End Sub

Sub 程序升级20180715()

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
        sheet1.Cells(3, i2) = "初始"
        sheet1.Cells(3, i2 + 1) = "即时一"
        sheet1.Cells(3, i2 + 2) = "即时二"
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


Sub 程序升级20180813()

    Dim sheet1 As Worksheet
    Dim tmClBgCol As Long       '总排名交锋等级数据所在列

    Set sheet1 = ActiveWorkbook.Sheets("综合数据")
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    
    
    '处理Ok30
    tmClBgCol = dataColDict.Item("OK30_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("BFW")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "Ok30-标识"
        sheet1.Cells(3, i2) = "初始"
        sheet1.Cells(3, i2 + 1) = "即时一"
        sheet1.Cells(3, i2 + 2) = "即时二"
        sheet1.Cells(4, i2) = "OK30_1"
        
        sheet1.Range(Cells(2, i2 + 3), Cells(2, i2 + 4)).Merge
        sheet1.Cells(2, i2 + 3) = "Ok30-比较"
        sheet1.Cells(3, i2 + 3) = "即时一"
        sheet1.Cells(3, i2 + 4) = "即时二"

        sheet1.Cells(4, i2 + 3) = "OK30_2"
        
        MsgBox ("OK30分析程序更新完毕！")
    Else
        MsgBox ("OK30分析程序已更新！")
    End If
    
    
    '处理BF1
    tmClBgCol = dataColDict.Item("BF1_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("OKBF1")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "BF1-标识"
        sheet1.Cells(3, i2) = "初始"
        sheet1.Cells(3, i2 + 1) = "即时一"
        sheet1.Cells(3, i2 + 2) = "即时二"
        sheet1.Cells(4, i2) = "BF1_1"
        
        sheet1.Range(Cells(2, i2 + 3), Cells(2, i2 + 5)).Merge
        sheet1.Cells(2, i2 + 3) = "BF1-比较"
        sheet1.Cells(3, i2 + 3) = "初始"
        sheet1.Cells(3, i2 + 4) = "即时一"
        sheet1.Cells(3, i2 + 5) = "即时二"

        sheet1.Cells(4, i2 + 3) = "BF1_2"
        
        MsgBox ("BF1分析程序更新完毕！")
    Else
        MsgBox ("BF1分析程序已更新！")
    End If
    
    
    '处理威廉
    tmClBgCol = dataColDict.Item("DATAW_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("DATAB")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert

        
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "威廉-初始"
        sheet1.Cells(3, i2) = "胜"
        sheet1.Cells(3, i2 + 1) = "平"
        sheet1.Cells(3, i2 + 2) = "负"
        sheet1.Cells(4, i2) = "DATAW_1"
        
        sheet1.Range(Cells(2, i2 + 3), Cells(2, i2 + 5)).Merge
        sheet1.Cells(2, i2 + 3) = "威廉-即时一"
        sheet1.Cells(3, i2 + 3) = "胜"
        sheet1.Cells(3, i2 + 4) = "平"
        sheet1.Cells(3, i2 + 5) = "负"
        sheet1.Cells(4, i2 + 3) = "DATAW_2"

        sheet1.Range(Cells(2, i2 + 6), Cells(2, i2 + 8)).Merge
        sheet1.Cells(2, i2 + 6) = "威廉-即时二"
        sheet1.Cells(3, i2 + 6) = "胜"
        sheet1.Cells(3, i2 + 7) = "平"
        sheet1.Cells(3, i2 + 8) = "负"
        sheet1.Cells(4, i2 + 6) = "DATAW_3"

        
        MsgBox ("威廉分析程序更新完毕！")
    Else
        MsgBox ("威廉分析程序已更新！")
    End If

    '处理模式8
    tmClBgCol = dataColDict.Item("SCHEMA8_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("DATAW")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "模式八并排"
        sheet1.Cells(3, i2) = "初始"
        sheet1.Cells(3, i2 + 1) = "即时一"
        sheet1.Cells(3, i2 + 2) = "即时二"
        sheet1.Cells(4, i2) = "SCHEMA8_1"
        
        
        MsgBox ("SCHEMA8并排  程序更新完毕！")
    Else
        MsgBox ("SCHEMA8并排 程序已更新！")
    End If
    
    '处理模式7
    tmClBgCol = dataColDict.Item("SCHEMA7_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("SCHEMA8")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "模式七并排"
        sheet1.Cells(3, i2) = "初始"
        sheet1.Cells(3, i2 + 1) = "即时一"
        sheet1.Cells(3, i2 + 2) = "即时二"
        sheet1.Cells(4, i2) = "SCHEMA7_1"
        
        
        MsgBox ("SCHEMA7并排  程序更新完毕！")
    Else
        MsgBox ("SCHEMA7并排 程序已更新！")
    End If
    
    
    '处理模式6——基于bet365盘口
    tmClBgCol = dataColDict.Item("SCHEMA6_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("SCHEMA7")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        
        
        sheet1.Cells(2, i2) = "模式6-Bet365"
        sheet1.Cells(3, i2) = "模式6-Bet365"
        sheet1.Cells(4, i2) = "SCHEMA6_1"
        
        
        MsgBox ("SCHEMA6-Bet365盘口  程序更新完毕！")
    Else
        MsgBox ("SCHEMA6-Bet365盘口 程序已更新！")
    End If
    
    
End Sub

Sub 程序升级20180818()

    Dim sheet1 As Worksheet
    Dim tmClBgCol As Long       '总排名交锋等级数据所在列

    Set sheet1 = ActiveWorkbook.Sheets("综合数据")
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    
    '找到SCHEMA模式一对应的列，将模式2和模式3定义SCHEMA2,SCHEMA3
    tmClBgCol = dataColDict.Item("SCHEMA1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("SCHEMA")
        sheet1.Cells(4, i2) = "SCHEMA1"
        sheet1.Cells(4, i2 + 1) = "SCHEMA2"
        sheet1.Cells(4, i2 + 2) = "SCHEMA3"
        Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    End If
    
    
    '处理模式4
    tmClBgCol = dataColDict.Item("SCHEMA4_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("SCHEMA5")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "模式四并排"
        sheet1.Cells(3, i2) = "初始"
        sheet1.Cells(3, i2 + 1) = "即时一"
        sheet1.Cells(3, i2 + 2) = "即时二"
        sheet1.Cells(4, i2) = "SCHEMA4_1"
        
        
        MsgBox ("SCHEMA4并排  程序更新完毕！")
    Else
        MsgBox ("SCHEMA4并排 程序已更新！")
    End If
    
    '处理模式2
    tmClBgCol = dataColDict.Item("SCHEMA2_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("SCHEMA3")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "模式二并排"
        sheet1.Cells(3, i2) = "初始"
        sheet1.Cells(3, i2 + 1) = "即时一"
        sheet1.Cells(3, i2 + 2) = "即时二"
        sheet1.Cells(4, i2) = "SCHEMA2_1"
        
        
        MsgBox ("SCHEMA2并排  程序更新完毕！")
    Else
        MsgBox ("SCHEMA2并排 程序已更新！")
    End If
        
    '处理模式1
    tmClBgCol = dataColDict.Item("SCHEMA1_1")

    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("SCHEMA2")
        '加入3列
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        sheet1.Cells.Columns(i2).Insert
        
        sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
        sheet1.Cells(2, i2) = "模式一并排"
        sheet1.Cells(3, i2) = "初始"
        sheet1.Cells(3, i2 + 1) = "即时一"
        sheet1.Cells(3, i2 + 2) = "即时二"
        sheet1.Cells(4, i2) = "SCHEMA1_1"
        
        
        MsgBox ("SCHEMA1并排  程序更新完毕！")
    Else
        MsgBox ("SCHEMA1并排 程序已更新！")
    End If
    
    
    
    
End Sub


Sub 程序升级20180827()

    Dim sheet1 As Worksheet
    Dim tmClBgCol As Long       '等处理数据所在列
    Dim colDesp
    Dim colIndex
    Dim insertColIdx

    Set sheet1 = ActiveWorkbook.Sheets("综合数据")
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    
    '数据按从后往前的顺序排列。这样有效的利用现有的字典中列的值。
    
    colDesp = Split("99家平均,澳门,Bet365", ",")
    colIndex = Split("99AVGRATIO,DATAM,DATAB", ",")
    insertColIdx = Split("CMPRATIO,DATAJ,DATAM", ",")
    
    If UBound(colDesp) <> UBound(colIndex) Or UBound(colIndex) <> UBound(insertColIdx) Then
        MsgBox ("升级失败！")
        Exit Sub
    End If
    
    For i = 0 To UBound(colIndex)
        tmClBgCol = dataColDict.Item(colIndex(i) & "_1")
    
        If tmClBgCol = 0 Then
            i2 = dataColDict.Item(insertColIdx(i))
            '加入3列
            sheet1.Cells.Columns(i2).Insert
            sheet1.Cells.Columns(i2).Insert
            sheet1.Cells.Columns(i2).Insert
            sheet1.Cells.Columns(i2).Insert
            sheet1.Cells.Columns(i2).Insert
            sheet1.Cells.Columns(i2).Insert
            sheet1.Cells.Columns(i2).Insert
            sheet1.Cells.Columns(i2).Insert
            sheet1.Cells.Columns(i2).Insert
    
            
            sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
            sheet1.Cells(2, i2) = colDesp(i) & "-初始"
            sheet1.Cells(3, i2) = "胜"
            sheet1.Cells(3, i2 + 1) = "平"
            sheet1.Cells(3, i2 + 2) = "负"
            sheet1.Cells(4, i2) = colIndex(i) & "_1"
            
            sheet1.Range(Cells(2, i2 + 3), Cells(2, i2 + 5)).Merge
            sheet1.Cells(2, i2 + 3) = colDesp(i) & "-即时一"
            sheet1.Cells(3, i2 + 3) = "胜"
            sheet1.Cells(3, i2 + 4) = "平"
            sheet1.Cells(3, i2 + 5) = "负"
            sheet1.Cells(4, i2 + 3) = colIndex(i) & "_2"
    
            sheet1.Range(Cells(2, i2 + 6), Cells(2, i2 + 8)).Merge
            sheet1.Cells(2, i2 + 6) = colDesp(i) & "-即时二"
            sheet1.Cells(3, i2 + 6) = "胜"
            sheet1.Cells(3, i2 + 7) = "平"
            sheet1.Cells(3, i2 + 8) = "负"
            sheet1.Cells(4, i2 + 6) = colIndex(i) & "_3"
            
            sheet1.Range(Cells(2, i2), Cells(3, i2 + 8)).Borders.LineStyle = 1
            MsgBox (colDesp(i) & "-分析程序更新完毕！")
        Else
            MsgBox (colDesp(i) & "-分析程序已更新！")
        End If
    Next
    
    
    '重新加载一次列字典，以更新前面追加的列信息
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    
    '一次性增加，按正序排列
    colDesp = Split("赔2,赔1", ",")
    colIndex = Split("LOSE2,LOSE1", ",")
    insertColIdx = Split("即时二,即时一,初始", ",")
    
    If UBound(colDesp) <> UBound(colIndex) Then
        MsgBox ("升级失败！")
        Exit Sub
    End If
    
    tmClBgCol = dataColDict.Item(colIndex(0) & "_1")
    
    
    If tmClBgCol = 0 Then
        i2 = dataColDict.Item("BF1")
        
        For i = 0 To 2    '初始、即时一、即时二
        
            For j = 0 To UBound(colIndex)
                '加入3列
                sheet1.Cells.Columns(i2).Insert
                sheet1.Cells.Columns(i2).Insert
                sheet1.Cells.Columns(i2).Insert
        
                
                sheet1.Range(Cells(2, i2), Cells(2, i2 + 2)).Merge
                sheet1.Cells(2, i2) = colDesp(j) & "-" & insertColIdx(i)
                sheet1.Cells(3, i2) = "胜"
                sheet1.Cells(3, i2 + 1) = "平"
                sheet1.Cells(3, i2 + 2) = "负"
                sheet1.Cells(4, i2) = colIndex(j) & "_" & (3 - i)
            Next
        Next
        MsgBox (colDesp(0) & "-分析程序更新完毕！")
    Else
        MsgBox (colDesp(0) & "-分析程序已更新！")
    End If

    
    
End Sub



Sub 程序升级20190907()

    Dim sheet1 As Worksheet
    Dim tmClBgCol As Long       '等处理数据所在列
    Dim colDesp
    Dim colIndex
    Dim insertColIdx

    Set sheet1 = ActiveWorkbook.Sheets("综合数据")
    Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    
    '数据按从后往前的顺序排列。这样有效的利用现有的字典中列的值。
    
     '删除99家平均，即时值二，即时值一、初始值
    colIndex = Split("99AVGRATIO_1,99AVGRATIO_2,99AVGRATIO_3", ",")
    For i = 0 To UBound(colIndex)
        
        i2 = dataColDict.Item(colIndex(i))
        If i2 > 0 Then
            sheet1.Cells.Columns(i2).Delete
            sheet1.Cells.Columns(i2).Delete
            sheet1.Cells.Columns(i2).Delete
            Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
        End If
        
    Next
    
    '删除竞彩比例四项
    
    i2 = dataColDict.Item("CMPRATIO")
    If i2 > 0 Then
        sheet1.Cells.Columns(i2).Delete
        sheet1.Cells.Columns(i2).Delete
        sheet1.Cells.Columns(i2).Delete
        sheet1.Cells.Columns(i2).Delete
        Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
    End If
   
    
    
    colDesp = Split("99家平均,澳门,Bet365", ",")
    colIndex = Split("DATAW,DATAW_1,DATAW_2,DATAW_3,DATAB,DATAB_1,DATAB_2,DATAB_3,DATAM,DATAM_1,DATAM_2,DATAM_3,DATAL,DATAE,LOSE1,LOSE2", ",")
    insertColIdx = Split("DATAW_1,DATAW_2,DATAW_3,DATAB,DATAB_1,DATAB_2,DATAB_3,DATAM,DATAM_1,DATAM_2,DATAM_3,DATAJ,DATAE,LOSE1,LOSE2,LOSE1_1", ",")
    colCnt = Split("4,3,3,3,4,3,3,3,4,3,3,3,4,4,4,4", ",")
    If UBound(colCnt) <> UBound(colIndex) Or UBound(colIndex) <> UBound(insertColIdx) Then
        MsgBox ("升级失败！")
        Exit Sub
    End If
    
    For i = 0 To UBound(colIndex)
         i2 = dataColDict.Item(colIndex(i))
    
        If i2 > 0 Then
            tmClBgCol = dataColDict.Item(insertColIdx(i))
            i3 = Int(colCnt(i))
            
            If tmClBgCol - i2 = i3 And i3 >= 3 Then '还没有增加列
            
                sheet1.Cells.Columns(i2 + 3).Insert
        
                sheet1.Cells(3, i2 + 3) = "返还率"
                If i3 = 3 Then   '
                    sheet1.Range(Cells(2, i2), Cells(2, i2 + 3)).Merge
                End If
                sheet1.Range(Cells(2, i2), Cells(3, i2 + 3)).Borders.LineStyle = 1
                Call 初始化一般字典(dataColDict, sheet1, 4, 0, 1, False)
            
                MsgBox (colIndex(i) & "-分析程序更新完毕！")
            End If
        Else
            MsgBox (colIndex(i) & "-分析程序已更新！")
        End If
    Next
    
    MsgBox ("全部更新完毕！")
End Sub

Sub 澳客网数据导入()
Dim begindate As Date
Dim enddate As Date

    begindate = DateAdd("d", -1, Date)
    enddate = DateAdd("d", 2, Date)
    
    Call 澳客网必发盈亏(begindate, enddate)
    'Call 澳客网胜负指数(begindate, enddate)
    'Call 澳客网盘口评测(begindate, enddate)
    'Call 澳客网凯利指数(begindate, enddate)
    'Call 澳客网数据载入
    MsgBox ("导入完毕！")
End Sub

Sub 测试积分数据导入()
    Call 初始化字典(leagueDict, "01赛事")
    Call 球探网赛事积分数据载入
End Sub
