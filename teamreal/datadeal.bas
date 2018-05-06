Attribute VB_Name = "datadeal"
Sub 数据更新()
    
Dim data()    '网站数据
Dim realData()   '球队实力值数据
Dim teamData()   '球队数据

Dim wkSheet As Worksheet   '综合数据
Dim realSheet As Worksheet   '球队实力
Dim row1, row2
Dim loc1, loc2


Dim dataDict  As Object
Dim realDict As Object

Dim vsId       '对阵ID
Dim datalen    '取实力数据的长度



Set wkSheet = ThisWorkbook.Sheets("综合数据")
Set realSheet = ThisWorkbook.Sheets("球队实力")


row1 = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
row2 = realSheet.UsedRange.Rows(realSheet.UsedRange.Rows.Count).row

'加载字典以便搜索
Call 初始化字典(Dict, "Param")
datalen = Dict.Item("TEAMREALDATALENGTH")



Call 初始化一般字典(dataDict, wkSheet, 5, 0, 2, True)
Call 初始化球队实力字典(realDict, "球队实力", datalen)


'数据加载部分
Call 加载数据到内存("球队信息", teamData, 2, 1)
Call 加载数据到内存("赛程积分", data, 2, 1)


For i = 2 To UBound(data)
    vsId = data(i, 3)     '对阵Id
    
    '处理实力数据
    If InStr(data(i, 9), "-") > 0 Then
        If Not dataDict.exists(vsId) Then    '综合数据中不存在当前对阵信息，移植数据
            row1 = row1 + 1
            Call 数据移植(data, i, wkSheet, row1)
        End If
        Call 实力值计算(data, i, realDict, realSheet)
    End If
Next

Set realDict = Nothing
Set dataDict = Nothing

    MsgBox ("数据更新完毕!")

End Sub


Sub 数据移植(data, row1, sheet1 As Worksheet, loc)
'********************************************************
'tempData 数据存放的数组
'row1   数据在数组中的行号
'sheet1：要存放的工作表
'loc：要存放的工作表的位置
'********************************************************
sheet1.Cells(loc, 1) = data(row1, 4)    '联赛 ID
sheet1.Cells(loc, 2) = data(row1, 26)    '联赛简称
sheet1.Cells(loc, 3) = data(row1, 1)    '赛季
sheet1.Cells(loc, 4) = data(row1, 2)    '轮次
sheet1.Cells(loc, 5) = data(row1, 3)    '对阵ID
sheet1.Cells(loc, 6) = data(row1, 6)    '日期

sheet1.Cells(loc, 7) = data(row1, 7)   '主队ID
sheet1.Cells(loc, 8) = data(row1, 27)   '主队名称
sheet1.Cells(loc, 9) = data(row1, 8)   '客队ID
sheet1.Cells(loc, 10) = data(row1, 28)   '客队名称
sheet1.Cells(loc, 11) = data(row1, 9)  '比分

sheet1.Cells(loc, 12) = data(row1, 10)   '半场比分
sheet1.Cells(loc, 13) = data(row1, 11)   '主队积分排名
sheet1.Cells(loc, 14) = data(row1, 12)   '客队积分排名
sheet1.Cells(loc, 15) = data(row1, 21)   '主队红牌次数
sheet1.Cells(loc, 16) = data(row1, 22)  '客队红牌次数

End Sub

Sub 实力值计算(data, i, dict1 As Object, sheet1 As Worksheet)
Dim priTeamId     '主队实力ID
Dim secTeamId     '客队实力Id
Dim teamId     '实力ID
Dim priReal    '主队实力值
Dim secReal    '客队实力值
Dim prePriReal   '原主队实力值
Dim preSecReal   '原客队实力值
Dim adjReal      '调整实力值
Dim score        '比分
Dim priScore     '主队得分
Dim secScore     '客队得分

Dim prelunCi, lunCi

Dim j, tempId


Dim row1
row1 = sheet1.UsedRange.Rows(sheet1.UsedRange.Rows.Count).row

    lunCi = data(i, 2)   '轮次
    prelunCi = data(i, 2) - 1   '前一轮次
    priTeamId = data(i, 4) & data(i, 7) & data(i, 1)  '主队实力Id
    secTeamId = data(i, 4) & data(i, 8) & data(i, 1)   '客队实力Id
    score = Split(data(i, 9), "-")
    priScore = score(0)
    secScore = score(1)
    
    
    'test begin
    'If (data(i, 7) = 273 Or data(i, 8) = 273) Then    '｛lunCi = 17 And｝
    '    MsgBox ("Test")
    'End If
    'test end
    
    
    If lunCi = 1 Then
        prePriReal = 100
        preSecReal = 100
    Else
        '取主队前一次的值
        teamId = priTeamId & prelunCi
        
        If dict1.exists(teamId) Then
            prePriReal = dict1.Item(teamId)
        Else  '不存在当前轮次的数据，取依次取前一次的积分，直到轮次小于1，update 2016.1.10
            j = prelunCi - 1
            
            Do While j >= 1
                tempId = priTeamId & j
                If dict1.exists(tempId) Then
                    prePriReal = dict1.Item(tempId)
                    Exit Do
                End If
                j = j - 1
            Loop
            If j = 0 Then
                prePriReal = 100
            End If
        End If
        
        '取客队前一轮次的值
        teamId = secTeamId & prelunCi
        If dict1.exists(teamId) Then
            preSecReal = dict1.Item(teamId)
        Else  '不存在当前轮次的数据，取依次取前一次的积分，直到轮次小于1,update 2016.1.10
            j = prelunCi - 1
            
            Do While j >= 1
                tempId = secTeamId & j
                If dict1.exists(tempId) Then
                    preSecReal = dict1.Item(tempId)
                    Exit Do
                End If
                j = j - 1
            Loop
            If j = 0 Then
                preSecReal = 100
            End If
        End If
        
    End If
    If priScore > secScore Then  '主胜
        adjReal = (prePriReal * 0.07 + preSecReal * 0.05) * 5 / 12
            
    ElseIf priScore < secScore Then '主负
        adjReal = -(prePriReal * 0.07 + preSecReal * 0.05) * 7 / 12
    Else
        adjReal = -(prePriReal * 0.07 + preSecReal * 0.05) / 12
    End If

    adjReal = Round(adjReal, 2)
    priReal = prePriReal + adjReal
    secReal = preSecReal - adjReal
    
    '字典中增加数据项
    priTeamId = priTeamId & lunCi
    
    secTeamId = secTeamId & lunCi
    If Not dict1.exists(priTeamId) Then
        dict1.Add priTeamId, priReal
        row1 = row1 + 1
        sheet1.Cells(row1, 1) = data(i, 4)     '联赛Id
        sheet1.Cells(row1, 2) = data(i, 26)     '联赛简称
        sheet1.Cells(row1, 3) = data(i, 7)     '球队Id
        sheet1.Cells(row1, 4) = data(i, 27)     '球队简称
        sheet1.Cells(row1, 5) = data(i, 1)     '赛季
        sheet1.Cells(row1, 6) = data(i, 2)     '轮次
        sheet1.Cells(row1, 7) = priReal     '实力值
    End If
    
    If Not dict1.exists(secTeamId) Then
        dict1.Add secTeamId, secReal
        row1 = row1 + 1
        sheet1.Cells(row1, 1) = data(i, 4)     '联赛Id
        sheet1.Cells(row1, 2) = data(i, 26)     '联赛简称
        sheet1.Cells(row1, 3) = data(i, 8)     '球队Id
        sheet1.Cells(row1, 4) = data(i, 28)     '球队简称
        sheet1.Cells(row1, 5) = data(i, 1)     '赛季
        sheet1.Cells(row1, 6) = data(i, 2)     '轮次
        sheet1.Cells(row1, 7) = secReal     '实力值
    End If

End Sub


Sub 视图刷新()
'刷新数据视图
    ThisWorkbook.Sheets("数据视图").PivotTables("数据透视表1").PivotCache.Refresh
End Sub
