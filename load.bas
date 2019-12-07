Attribute VB_Name = "load"
Option Explicit


Sub LoadDataToArray(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称
'minMove：最小移动指针
'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId
Dim dataArr
Dim str1 '日期
Dim vsId    '对阵的ID
Dim colCnt '列数

colCnt = 15

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
ReDim dataArr(cnt - 1, colCnt)
'将数据取到内存数组中
i = 2
Loc = 1
Do While i <= cnt
    If wkSheet.Cells(i, 2) <> "" Then
        '判断是否要求选择的联赛比赛，通过字典进行判断，加快速度
        itemId = wkSheet.Cells(i, 2)
        If leagueDict.exists(itemId) Then
            '获得下级链接的地址
            vsId = CDate(wkSheet.Cells(i, 3).Text)
            dataArr(Loc, 0) = wkSheet.Cells(i, 13)               '赛事ID
            dataArr(Loc, 1) = wkSheet.Cells(i, 2) '联赛
            dataArr(Loc, 2) = CDate(Year(vsId) & "/" & Month(vsId) & "/" & Day(vsId))    'CDate(Left(vsId, 10)) '日期
            dataArr(Loc, 3) = CDate(Right(vsId, 8)) '时间
            dataArr(Loc, 4) = wkSheet.Cells(i, 4)  '主队
            dataArr(Loc, 5) = wkSheet.Cells(i, 12)  '客队
            dataArr(Loc, 6) = wkSheet.Cells(i, 4) + " VS " + wkSheet.Cells(i, 12)  '对阵
            dataArr(Loc, 7) = wkSheet.Cells(i, 8) '主胜率(初始值)
            dataArr(Loc, 8) = wkSheet.Cells(i, 9) '和率(初始值)
            dataArr(Loc, 9) = wkSheet.Cells(i, 10) '客胜率(初始值)
            dataArr(Loc, 10) = wkSheet.Cells(i, 11)   '返还率（初始值）
            dataArr(Loc, 11) = wkSheet.Cells(i, 17) '主胜率（即时值)
            dataArr(Loc, 12) = wkSheet.Cells(i, 18) '和率（即时值)
            dataArr(Loc, 13) = wkSheet.Cells(i, 19) '客胜率（即时值)
            dataArr(Loc, 14) = wkSheet.Cells(i, 20)   '返还率（即时值)
            dataArr(Loc, 15) = wkSheet.Cells(i, 1) '比分
            Loc = Loc + 1
        End If
        '指针移位
        i = i + 1
    Else
        i = i + 1
    End If
Loop
ReDim datas(Loc - 1, colCnt)
'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To colCnt
        datas(i, j) = dataArr(i, j)
    Next
Next
Set wkSheet = Nothing

End Sub



Sub BF数据载入(datas, sheetName As String)
'------------------------------------------------------------
'dataBF:数据输出的数组
'dataW:要查找的数据根据data(,0)的ID号去链接新的网址数据
'------------------------------------------------------------
Dim rowNo
Dim col
Dim i, j
Dim vsId
Dim bfData() As Double

Dim dataArr
Dim Loc
Dim wkSheet As Worksheet

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
rowNo = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
ReDim dataArr(rowNo - 1, 32)

   'id号，四个指标（赔1、赔2、bf1、bf3），前2个指标6个数据，后2个指标各有8个
    '1-4  :赔1：初始值（胜平负）  5-8  :赔1：即时值（胜平负）
    '9-12：赔2：初始值（胜平负）  13-16：赔2：即时值（胜平负）
    '17-20：bf1:初始值（胜、平、负、返还率）  21——24：bf1:即时值（胜、平、负、返还率）
    '25-28：bf2:初始值（胜、平、负、返还率）  29——32：bf2:即时值（胜、平、负、返还率）
Loc = 1
For i = 2 To rowNo
    If wkSheet.Cells(i, 1) <> "" Then
        For j = 1 To 33
            If wkSheet.Cells(i, j) <> 0 Then
                dataArr(Loc, j - 1) = wkSheet.Cells(i, j)
            End If
        Next
        dataArr(Loc, 28) = dataArr(Loc, 20)    '2019.9.7 将beffair的返还率2的数据填入BF1的返还率（初始值)
        dataArr(Loc, 32) = dataArr(Loc, 24)    '2019.9.7 将beffair的返还率2的数据填入BF1的返还率(即时值）。
        Loc = Loc + 1
    End If
Next

ReDim datas(Loc - 1, 32)
'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To 32
        datas(i, j) = dataArr(i, j)
    Next
Next
'清除缓存
Set wkSheet = Nothing


End Sub

Sub 球探网联赛积分载入(datas, sheetName As String)
'------------------------------------------------------------
'dataBF:数据输出的数组
'dataW:要查找的数据根据data(,0)的ID号去链接新的网址数据
'------------------------------------------------------------
Dim rowNo
Dim col
Dim i, j
Dim vsId
Dim bfData() As Double

Dim dataArr
Dim Loc
Dim wkSheet As Worksheet

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
rowNo = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
ReDim dataArr(rowNo - 1, 104)

Loc = 1
For i = 2 To rowNo
    If wkSheet.Cells(i, 1) <> "" Then
        dataArr(Loc, 0) = wkSheet.Cells(i, 1) '保存id
        For j = 2 To 89
            If j Mod 11 = 4 Then
                dataArr(Loc, j - 1) = wkSheet.Cells(i, j)
                dataArr(Loc, j) = wkSheet.Cells(i, j + 1)
                dataArr(Loc, j + 1) = wkSheet.Cells(i, j + 2)
                dataArr(Loc, j + 2) = wkSheet.Cells(i, j + 6)
                dataArr(Loc, j + 3) = wkSheet.Cells(i, j + 7)
            End If
        Next
        
        '欧亚转盘数据
        
        For j = 90 To 105
            dataArr(Loc, j - 1) = wkSheet.Cells(i, j)
        Next
        Loc = Loc + 1
    End If
Next

ReDim datas(Loc - 1, 104)
'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To 104
            datas(i, j) = dataArr(i, j)
    Next
Next
'清除缓存
Set wkSheet = Nothing


End Sub




Sub 澳客网胜负指数载入(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称

'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId
Dim dataArr
Dim currDate


Dim leagueData()
Dim league

'取对应关系数据
Call loadLeagueData(leagueData)

Call 初始化字典(teamUniDict, "02球队", 2, 1, 2)

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.Cells(1, 1)
ReDim dataArr(cnt - 1, 14)
'将数据取到内存数组中
Loc = 1
For i = 1 To cnt - 1
    If wkSheet.Cells(i, 2) <> "" Then
        '判断是否要求选择的联赛比赛，通过字典进行判断，加快速度
        league = wkSheet.Cells(i + 1, 4) '赛事
        league = UniformLeague(leagueData, league, 2)
        If leagueDict.exists(league) Then
            dataArr(Loc, 1) = league '联赛
            dataArr(Loc, 2) = wkSheet.Cells(i + 1, 5) '日期
            dataArr(Loc, 3) = CDate(wkSheet.Cells(i + 1, 6))  '时间
            dataArr(Loc, 4) = UniformTeam(teamUniDict, Trim(wkSheet.Cells(i + 1, 7))) '主队
            dataArr(Loc, 5) = UniformTeam(teamUniDict, Trim(wkSheet.Cells(i + 1, 9)))  '客队
            dataArr(Loc, 6) = Trim(wkSheet.Cells(i + 1, 8))  '对阵
            dataArr(Loc, 7) = wkSheet.Cells(i + 1, 10) '主胜率(初始值)
            dataArr(Loc, 8) = wkSheet.Cells(i + 1, 11) '和率(初始值)
            dataArr(Loc, 9) = wkSheet.Cells(i + 1, 12) '客胜率(初始值)
            dataArr(Loc, 10) = CDbl(Left(wkSheet.Cells(i + 1, 13), 4)) '主胜率（即时值)
            dataArr(Loc, 11) = CDbl(Left(wkSheet.Cells(i + 1, 14), 4)) '和率（即时值)
            dataArr(Loc, 12) = CDbl(Left(wkSheet.Cells(i + 1, 15), 4)) '客胜率（即时值)
            dataArr(Loc, 13) = wkSheet.Cells(i + 1, 1)  '当期期数
            dataArr(Loc, 14) = wkSheet.Cells(i + 1, 3)  '当期编号
            Loc = Loc + 1
        End If
    End If
Next

ReDim datas(Loc - 1, 14)

'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 1 To 14
        datas(i, j) = dataArr(i, j)
    Next
Next

Set wkSheet = Nothing
End Sub


Sub 澳客网凯利方差载入(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称
'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId, itemid1
Dim team
Dim teamInfo
Dim dataArr




Dim leagueData()
Dim league
Dim tempData



'取对应关系数据
Call loadLeagueData(leagueData)
Call 初始化字典(teamUniDict, "02球队", 2, 1, 2)


Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.Cells(1, 1)
ReDim dataArr(cnt - 1, 21)   '6个基本数据，4*6个指标
'将数据取到内存数组中
Loc = 1
For i = 1 To cnt - 1
        
    team = wkSheet.Cells(i + 1, 4)
    team = UniformLeague(leagueData, team, 2)
    If leagueDict.exists(team) Then
        dataArr(Loc, 1) = team '联赛
        dataArr(Loc, 2) = wkSheet.Cells(i + 1, 5) '日期
        dataArr(Loc, 3) = CDate(wkSheet.Cells(i + 1, 6)) '时间
        dataArr(Loc, 4) = UniformTeam(teamUniDict, Trim(wkSheet.Cells(i + 1, 7)))  '主队
        dataArr(Loc, 5) = UniformTeam(teamUniDict, Trim(wkSheet.Cells(i + 1, 9)))  '客队
        dataArr(Loc, 6) = Trim(wkSheet.Cells(i + 1, 8))  '对阵
        
        '威廉希尔
        dataArr(Loc, 7) = wkSheet.Cells(i + 1, 10) '主胜率(初始值)
        dataArr(Loc, 8) = wkSheet.Cells(i + 1, 11) '和率(初始值)
        dataArr(Loc, 9) = wkSheet.Cells(i + 1, 12) '客胜率(初始值)
        dataArr(Loc, 10) = wkSheet.Cells(i + 1, 13) '赔付率
        
        
        'Bet365
        dataArr(Loc, 11) = wkSheet.Cells(i + 1, 14) '主胜率(初始值)
        dataArr(Loc, 12) = wkSheet.Cells(i + 1, 15) '和率(初始值)
        dataArr(Loc, 13) = wkSheet.Cells(i + 1, 16) '客胜率(初始值)
        dataArr(Loc, 14) = wkSheet.Cells(i + 1, 17) '赔付率

        
        '澳门网
        dataArr(Loc, 15) = wkSheet.Cells(i + 1, 18) '主胜率(初始值)
        dataArr(Loc, 16) = wkSheet.Cells(i + 1, 19) '和率(初始值)
        dataArr(Loc, 17) = wkSheet.Cells(i + 1, 20) '客胜率(初始值)
        dataArr(Loc, 18) = wkSheet.Cells(i + 1, 21) '赔付论
        
        '凯利方差
        dataArr(Loc, 19) = wkSheet.Cells(i + 1, 22) '主胜率(初始值)
        dataArr(Loc, 20) = wkSheet.Cells(i + 1, 23) '和率(初始值)
        dataArr(Loc, 21) = wkSheet.Cells(i + 1, 24) '客胜率(初始值)
        Loc = Loc + 1
    End If
Next
ReDim datas(Loc - 1, 21)

'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 1 To 21
        datas(i, j) = dataArr(i, j)
    Next
Next
Set wkSheet = Nothing
End Sub


Sub 澳客网必发盈亏载入(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称
'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i, j

Dim team   '联赛

Dim leagueData()
Dim dataArr()
Dim league
Dim amount1 As Double   '胜成交金额
Dim amount2 As Double   '平成交金额
Dim amount3 As Double   '负成交金额

'取对应关系数据
Call loadLeagueData(leagueData)
Call 初始化字典(teamUniDict, "02球队", 2, 1, 2)

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.Cells(1, 1)
'2015.9.13添加9列
ReDim dataArr(cnt - 1, 57)

'将数据取到内存数组中
Loc = 1
For i = 1 To cnt - 1
    team = wkSheet.Cells(i + 1, 4)
    team = UniformLeague(leagueData, team, 2)
    If leagueDict.exists(team) Then
        dataArr(Loc, 1) = team '联赛
        dataArr(Loc, 2) = wkSheet.Cells(i + 1, 5) '日期
        dataArr(Loc, 3) = CDate(wkSheet.Cells(i + 1, 6)) '时间
        dataArr(Loc, 4) = UniformTeam(teamUniDict, Trim(wkSheet.Cells(i + 1, 7)))  '主队
        dataArr(Loc, 5) = UniformTeam(teamUniDict, Trim(wkSheet.Cells(i + 1, 9)))  '客队
        dataArr(Loc, 6) = Trim(wkSheet.Cells(i + 1, 8))  '比分：未开始的为“VS”，结束的为实际比分
        dataArr(Loc, 7) = wkSheet.Cells(i + 1, 1)  '当期期数
        dataArr(Loc, 8) = wkSheet.Cells(i + 1, 3)  '当期编号
        
        
        For j = 10 To 48
            '----必发 OKBF1
            dataArr(Loc, j) = wkSheet.Cells(i + 1, j) '主胜率(初始值)
        Next
        
        '买家挂牌比例数据
        If dataArr(Loc, 10) <> "" Then amount1 = CDbl(dataArr(Loc, 10)) Else amount1 = 0
        If dataArr(Loc, 11) <> "" Then amount2 = CDbl(dataArr(Loc, 11)) Else amount2 = 0
        If dataArr(Loc, 12) <> "" Then amount3 = CDbl(dataArr(Loc, 12)) Else amount3 = 0
        
        If amount1 + amount2 + amount3 > 0 Then
            dataArr(Loc, 49) = Round(amount1 / (amount1 + amount2 + amount3), 4)
            dataArr(Loc, 50) = Round(amount2 / (amount1 + amount2 + amount3), 4)
            dataArr(Loc, 51) = Round(amount3 / (amount1 + amount2 + amount3), 4)
        End If
        
        
        '卖家挂牌比例数据
        If dataArr(Loc, 16) <> "" Then amount1 = CDbl(dataArr(Loc, 16)) Else amount1 = 0
        If dataArr(Loc, 17) <> "" Then amount2 = CDbl(dataArr(Loc, 17)) Else amount2 = 0
        If dataArr(Loc, 18) <> "" Then amount3 = CDbl(dataArr(Loc, 18)) Else amount3 = 0
        
        If amount1 + amount2 + amount3 > 0 Then
            dataArr(Loc, 52) = Round(amount1 / (amount1 + amount2 + amount3), 4)
            dataArr(Loc, 53) = Round(amount2 / (amount1 + amount2 + amount3), 4)
            dataArr(Loc, 54) = Round(amount3 / (amount1 + amount2 + amount3), 4)
        End If
        
        '必发.99平均赔比较
        If dataArr(Loc, 31) <> "" Then amount1 = CDbl(dataArr(Loc, 31)) * 0.95 - CDbl(dataArr(Loc, 37)) Else amount1 = 0
        If dataArr(Loc, 32) <> "" Then amount2 = CDbl(dataArr(Loc, 32)) * 0.95 - CDbl(dataArr(Loc, 38)) Else amount2 = 0
        If dataArr(Loc, 33) <> "" Then amount3 = CDbl(dataArr(Loc, 33)) * 0.95 - CDbl(dataArr(Loc, 39)) Else amount3 = 0
        
        dataArr(Loc, 55) = Round(amount1, 4)
        dataArr(Loc, 56) = Round(amount2, 4)
        dataArr(Loc, 57) = Round(amount3, 4)

        Loc = Loc + 1
    End If

Next


ReDim datas(Loc - 1, 57)

'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 1 To 57
        datas(i, j) = dataArr(i, j)
    Next
Next

Set wkSheet = Nothing
End Sub




Sub 澳客网盘口评测载入(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称
'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i, j

Dim team   '联赛

Dim leagueData()
Dim dataArr()
Dim league

'取对应关系数据
Call loadLeagueData(leagueData)
Call 初始化字典(teamUniDict, "02球队", 2, 1, 2)


Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.Cells(1, 1)
ReDim dataArr(cnt - 1, 49)
'将数据取到内存数组中
Loc = 1
For i = 1 To cnt - 1
    team = wkSheet.Cells(i + 1, 4)
    team = UniformLeague(leagueData, team, 2)
    If leagueDict.exists(team) Then
        dataArr(Loc, 1) = team '联赛
        dataArr(Loc, 2) = CDate(wkSheet.Cells(i + 1, 5)) '日期
        dataArr(Loc, 3) = CDate(wkSheet.Cells(i + 1, 6)) '时间
        dataArr(Loc, 4) = UniformTeam(teamUniDict, Trim(wkSheet.Cells(i + 1, 7)))  '主队
        dataArr(Loc, 5) = UniformTeam(teamUniDict, Trim(wkSheet.Cells(i + 1, 9)))  '客队
        dataArr(Loc, 6) = Trim(wkSheet.Cells(i + 1, 8))  '比分：未开始的为“VS”，结束的为实际比分
        
        
        For j = 10 To 49
            If Trim(wkSheet.Cells(i + 1, j)) <> "-" Then
                dataArr(Loc, j) = Trim(wkSheet.Cells(i + 1, j)) '主胜率(初始值)
            Else
                dataArr(Loc, j) = ""
            End If
        Next
        Loc = Loc + 1
    End If
Next


ReDim datas(Loc - 1, 49)

'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 1 To 49
        datas(i, j) = dataArr(i, j)
    Next
Next

Set wkSheet = Nothing
End Sub


Sub 综合数据载入内存(datas, sheetName As String, Optional recCnt As Long = 500, Optional bgCol As Long = 5)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称
'recCnt:载入的记录条数，默认为500，>0,表示具体条数，0：表示全部载入
'bgCol：开始读取的数据行，默认从第一个数据开始（Sheet中的第四行）
'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId
Dim dataArr


Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
ReDim dataArr(cnt - 1, 8)
'将数据取到内存数组中

'确定取数的范围，为0则取全部
If recCnt = 0 Then
    i = bgCol
ElseIf cnt - recCnt + 1 < bgCol Then
    i = bgCol
Else
    i = cnt - recCnt + 1
End If
If i > cnt Then
    MsgBox ("开始行大于页面中有数据的行号！")
    Exit Sub
End If

Loc = 1
Do While i <= cnt
    If wkSheet.Cells(i, 1) <> "" Then
        dataArr(Loc, 0) = i     '保存数据在sheet中对应的行号
        dataArr(Loc, 1) = wkSheet.Cells(i, 1) '日期
        dataArr(Loc, 2) = CDate(wkSheet.Cells(i, 2)) '时间
        dataArr(Loc, 3) = wkSheet.Cells(i, 3) '主队
        dataArr(Loc, 4) = wkSheet.Cells(i, 4)  '客队
        dataArr(Loc, 5) = wkSheet.Cells(i, 5)  '对阵
        dataArr(Loc, 6) = wkSheet.Cells(i, 6)  '数据类型
        dataArr(Loc, 7) = wkSheet.Cells(i, 7) '联赛
        dataArr(Loc, 8) = wkSheet.Cells(i, 9) '球探网赛事ID
        
        '指针移位
        Loc = Loc + 1
    End If
    i = i + 1
Loop
ReDim datas(Loc - 1, 8)
'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To 8
        datas(i, j) = dataArr(i, j)
    Next
Next
Set wkSheet = Nothing
End Sub


Sub 配置数据载入(dataArr, sheetName As String)
'------------------------------------------------------------------------
'将配置信息载入数组
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称
'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim col As Integer
Dim Loc As Long
Dim i As Long
Dim j As Long


Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
col = wkSheet.UsedRange.Columns(wkSheet.UsedRange.Columns.Count).Column
ReDim dataArr(cnt, col)
'将数据取到内存数组中
For i = 2 To cnt
    For j = 1 To col
        dataArr(i, j) = wkSheet.Cells(i, j)
    Next j
Next i
Set wkSheet = Nothing
End Sub


Sub BF载入赔率(datas, sheetName As String, userType As String)
'------------------------------------------------------------
'dataBF:数据输出的数组
'sheetName:Sheet页名称
'userType：对应的赔率类型，"B":bet365,"M":澳门，"L":立博，"E":易胜博
'------------------------------------------------------------
Dim rowNo
Dim col
Dim i, j
Dim vsId
Dim bfData() As Double

Dim dataArr
Dim Loc
Dim wkSheet As Worksheet
Dim bgCol As Integer

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
rowNo = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
ReDim dataArr(rowNo - 1, 14)

If userType <> "B" And userType <> "M" And userType <> "L" And userType <> "E" Then
    MsgBox ("从" + sheetName + "载入数据出错！")
    Exit Sub
End If

Select Case userType
    Case "B": bgCol = 34
    Case "M": bgCol = 42
    Case "L": bgCol = 50
    Case "E": bgCol = 58
End Select

Loc = 1
For i = 2 To rowNo
    If wkSheet.Cells(i, 1) <> "" Then
        dataArr(Loc, 0) = wkSheet.Cells(i, 1)               '赛事ID
        'dataArr(loc, 1) = wkSheet.Cells(i, 2) '联赛
        'dataArr(loc, 2) = CDate(Left(vsId, 10)) '日期
        'dataArr(loc, 3) = CDate(Right(vsId, 8)) '时间
        'dataArr(loc, 4) = wkSheet.Cells(i, 4)  '主队
        'dataArr(loc, 5) = wkSheet.Cells(i, 12)  '客队
        'dataArr(loc, 6) = wkSheet.Cells(i, 4) + " VS " + wkSheet.Cells(i, 12)  '对阵
        
        If wkSheet.Cells(i, bgCol) <> 0 Then
            dataArr(Loc, 7) = wkSheet.Cells(i, bgCol) '主胜率(初始值)
        End If
        If wkSheet.Cells(i, bgCol + 1) <> 0 Then
            dataArr(Loc, 8) = wkSheet.Cells(i, bgCol + 1) '和率(初始值)
        End If
        If wkSheet.Cells(i, bgCol + 2) <> 0 Then
            dataArr(Loc, 9) = wkSheet.Cells(i, bgCol + 2) '客胜率(初始值)
        End If
        If wkSheet.Cells(i, bgCol + 2) <> 0 Then
            dataArr(Loc, 10) = wkSheet.Cells(i, bgCol + 3) '返还率(初始值)
        End If
        
        If wkSheet.Cells(i, bgCol + 4) <> 0 Then
            dataArr(Loc, 11) = wkSheet.Cells(i, bgCol + 4) '主胜率（即时值)
        End If
        If wkSheet.Cells(i, bgCol + 5) <> 0 Then
            dataArr(Loc, 12) = wkSheet.Cells(i, bgCol + 5) '和率（即时值)
        End If
        If wkSheet.Cells(i, bgCol + 6) <> 0 Then
            dataArr(Loc, 13) = wkSheet.Cells(i, bgCol + 6) '客胜率（即时值)
        End If
        If wkSheet.Cells(i, bgCol + 2) <> 0 Then
            dataArr(Loc, 14) = wkSheet.Cells(i, bgCol + 7) '返还率(即时值)
        End If
        Loc = Loc + 1
    End If
Next

ReDim datas(Loc - 1, 14)
'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To 14
        datas(i, j) = dataArr(i, j)
    Next
Next
'清除缓存
Set wkSheet = Nothing

End Sub



Sub 澳客网数据载入()
Dim IE As Object
'Dim IE As MSXML2.XMLHTTP

Dim BFarr() As String       '必发数据
Dim SFarr() As String       '胜负数据
Dim KLarr() As String       '凯利数据

Dim tt
Dim doc As Object
Dim k As Long                '期数循环
Dim i As Long                '各期中球赛循环指针
Dim j As Integer            '球赛各公司的数据指针

Dim t1 As Integer
Dim t2 As Integer
Dim lotteryNo
Dim urlStr As String
Dim divObjects As Object     'DIV objects数组
Dim divObj As Object         '单个DIV对象
Dim node   As Object          '单个DIV的子对象
Dim BasicInfo As Object       '球赛基本信息
Dim Klinfo  As Object         '球赛的凯利指数
Dim lotterySel As Object      '期数选择按钮

Dim lotteryArr()
Dim len1  As Integer                   '可查询的期数
Dim Loc  As Integer                    '数据数组行指针

Dim wkSheet As Worksheet
Dim dataSheet As Worksheet


Set dataSheet = ActiveWorkbook.Sheets("综合数据")



Set IE = UserForm1.WebBrowser1
urlStr = "http://www.okooo.com/zucai/shuju/peilv/"    '+ lotteryNo + "/"

Set IE = UserForm1.WebBrowser1

With IE
  .Navigate urlStr '网址
  Do Until .ReadyState = 4
    DoEvents
  Loop
  Set doc = .document
End With
Application.ScreenUpdating = False

'针对出现500的错误，增加此段判断，add by ljqu 2016.9.4
If doc.nameProp = "HTTP 500 INternal Server Error" Then
    MsgBox ("http://www.okooo.com/zucai/shuju/peilv/不可访问")
    Exit Sub
End If

'取出当前对象
Set lotterySel = doc.body.All.tags("Select")
len1 = lotterySel(0).Length
ReDim lotteryArr(len1 - 1)
For i = 0 To len1 - 1
    lotteryArr(i) = lotterySel(0)(i).Value
Next

'保存当期的期数
dataSheet.Cells(1, 9) = lotterySel(0).Value
Application.CommandBars("彩票分析").Controls(6).Caption = "查看【" + CStr(lotterySel(0).Value) + "】期"


'-----------------------------------处理澳客网（胜负指数）开始----------------------------------------
'初始化接收数据的数组
'体彩期数，标识，编号，联赛，日期，时间，主队，比分，客队，初始(主胜、平局、客胜）、即时（主胜、平局、客胜）
ReDim SFarr(14 * len1, 14)
Loc = 1
'循环取所有期数的数据
For k = 0 To len1 - 1
    lotteryNo = lotteryArr(len1 - 1 - k)
    urlStr = "http://www.okooo.com/zucai/shuju/zhishu/" + lotteryNo + "/"
    With IE
      .Navigate urlStr '网址
      Do Until .ReadyState = 4
        DoEvents
      Loop
      Set doc = .document
    End With
    Application.ScreenUpdating = False
    
    
    Set divObjects = doc.body.All.tags("TABLE")
    For i = 0 To divObjects.Length - 1
        Set divObj = divObjects(i)
        If divObj.className = "magazine_table" Then     '输出的数据
              Set Klinfo = divObj.Rows
              For j = 1 To Klinfo.Length - 1            '第一行为标题，忽略
              
                SFarr(Loc, 0) = lotteryNo                 '期数
                SFarr(Loc, 1) = Klinfo(j).Cells(0).innerText + Klinfo(j).Cells(1).innerText + Klinfo(j).Cells(2).innerText 'Klinfo(j).innerText                   '标识
                SFarr(Loc, 2) = Klinfo(j).Cells(0).innerText     '编号
                SFarr(Loc, 3) = Klinfo(j).Cells(1).innerText     '联赛
                tt = Split(Klinfo(j).Cells(2).innerText, " ")
                SFarr(Loc, 4) = tt(0)     '日期
                SFarr(Loc, 5) = tt(1)     '时间
                
                
                SFarr(Loc, 6) = Klinfo(j).Cells(3).innerText     '主队
                SFarr(Loc, 7) = Klinfo(j).Cells(4).innerText     '比分
                SFarr(Loc, 8) = Klinfo(j).Cells(5).innerText     '客队
                
                '初始值
                SFarr(Loc, 9) = Klinfo(j).Cells(6).innerText                       '主胜
                SFarr(Loc, 10) = Klinfo(j).Cells(7).innerText                     '平局
                SFarr(Loc, 11) = Klinfo(j).Cells(8).innerText                     '客胜
                
                '即时值
                SFarr(Loc, 12) = Klinfo(j).Cells(9).innerText                       '主胜
                SFarr(Loc, 13) = Klinfo(j).Cells(10).innerText                     '平局
                SFarr(Loc, 14) = Klinfo(j).Cells(11).innerText                     '客胜
                Loc = Loc + 1
              Next
        End If
    Next
    '清除本次的变量
    Set Klinfo = Nothing
    Set BasicInfo = Nothing
    Set divObj = Nothing
    Set divObjects = Nothing
    'Set IE = Nothing

Next



Set wkSheet = ActiveWorkbook.Sheets("澳客网期数")
wkSheet.Cells.Clear


wkSheet.Cells(1, 1) = Loc
wkSheet.Cells(1, 2) = "标识"
wkSheet.Cells(1, 3) = "本期编号"
wkSheet.Cells(1, 4) = "联赛"
wkSheet.Cells(1, 5) = "日期"
wkSheet.Cells(1, 6) = "时间"
wkSheet.Cells(1, 7) = "主队"
wkSheet.Cells(1, 8) = "比分"
wkSheet.Cells(1, 9) = "客队"
wkSheet.Cells(1, 10) = "初始值-主胜"
wkSheet.Cells(1, 11) = "平局"
wkSheet.Cells(1, 12) = "客胜"
wkSheet.Cells(1, 13) = "即时值-主胜"
wkSheet.Cells(1, 14) = "平局"
wkSheet.Cells(1, 15) = "客胜"

wkSheet.Columns("C:C").NumberFormatLocal = "@"
wkSheet.Columns("E:E").NumberFormatLocal = "yyyy/m/d"
wkSheet.Columns("H:H").NumberFormatLocal = "@"

For i = 1 To Loc - 1
   For j = 0 To 14
        wkSheet.Cells(i + 1, j + 1) = SFarr(i, j)
   Next
Next i

Set wkSheet = Nothing

'-----------------------------------处理澳客网（胜负指数）结束----------------------------------------
Set IE = Nothing

End Sub



Sub 加载中国竞彩网数据(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称

'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId
Dim dataArr
Dim currDate


Dim colCnt As Long   '要加载的数据列总数

Dim leagueData()
Dim league

'取对应关系数据
Call loadLeagueData(leagueData)

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row


colCnt = 21
ReDim dataArr(cnt - 1, colCnt)
'将数据取到内存数组中
Loc = 1

For i = 1 To cnt - 1
    If wkSheet.Cells(i, 1) <> "" Then
        '判断是否要求选择的联赛比赛，通过字典进行判断，加快速度
        league = wkSheet.Cells(i + 1, 4) '赛事
        league = UniformLeague(leagueData, league, 4)
        If leagueDict.exists(league) Then
            dataArr(Loc, 0) = wkSheet.Cells(i + 1, 1) 'ID
            dataArr(Loc, 1) = CDate(wkSheet.Cells(i + 1, 2)) '日期
            dataArr(Loc, 2) = wkSheet.Cells(i + 1, 3)  '编号
            dataArr(Loc, 3) = league  '赛事
            dataArr(Loc, 4) = Trim(wkSheet.Cells(i + 1, 5))  '主队
            dataArr(Loc, 5) = Trim(wkSheet.Cells(i + 1, 6))  '客队
            
            
            For j = 6 To colCnt
                dataArr(Loc, j) = wkSheet.Cells(i + 1, j + 1)
            Next
            
            
            'dataArr(loc, 6) = wkSheet.Cells(i + 1, 7) '主胜(初始值)
            'dataArr(loc, 7) = wkSheet.Cells(i + 1, 8) '平(初始值)
            'dataArr(loc, 8) = wkSheet.Cells(i + 1, 9) '主负率(初始值)
            'dataArr(loc, 9) = wkSheet.Cells(i + 1, 10) '主胜（即时值)
            'dataArr(loc, 10) = wkSheet.Cells(i + 1, 11) '平（即时值)
            'dataArr(loc, 11) = wkSheet.Cells(i + 1, 12) '主负率(即时值)
            
            
            'dataArr(loc, 12) = wkSheet.Cells(i + 1, 13) '
            'dataArr(loc, 13) = wkSheet.Cells(i + 1, 14) '
            'dataArr(loc, 14) = wkSheet.Cells(i + 1, 15) '
            'dataArr(loc, 15) = wkSheet.Cells(i + 1, 16) '
            'dataArr(loc, 16) = wkSheet.Cells(i + 1, 17) '
            'dataArr(loc, 17) = wkSheet.Cells(i + 1, 18) '
            'dataArr(loc, 18) = wkSheet.Cells(i + 1, 19) '
            'dataArr(loc, 19) = wkSheet.Cells(i + 1, 20) '让球主胜投资比例
            'dataArr(loc, 20) = wkSheet.Cells(i + 1, 21) '让球平投资比例
            'dataArr(loc, 21) = wkSheet.Cells(i + 1, 22) '让球主负投资比例
            
            Loc = Loc + 1
        End If
    End If
Next

ReDim datas(Loc - 1, colCnt)

'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To colCnt
        datas(i, j) = dataArr(i, j)
    Next
Next

Set wkSheet = Nothing
End Sub



Sub 数据载入通用程序(datas, wkSheet As Worksheet, Optional recCnt As Long = 500, Optional bgCol As Long = 4)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称
'recCnt:载入的记录条数，默认为500，>0,表示具体条数，0：表示全部载入
'bgCol：开始读取的数据行，默认从第一个数据开始（Sheet中的第四行）
'------------------------------------------------------------------------
Dim row As Long
Dim col As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId
Dim dataArr


row = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
col = wkSheet.UsedRange.Columns(wkSheet.UsedRange.Columns.Count).Column
ReDim dataArr(row, col)
'将数据取到内存数组中

'确定取数的范围，为0则取全部
If recCnt = 0 Then
    i = bgCol
ElseIf row - recCnt + 1 < bgCol Then
    i = bgCol
Else
    i = row - recCnt + 1
End If
If i > row Then
    MsgBox ("开始行大于页面中有数据的行号！")
    Exit Sub
End If

Loc = 1
Do While i <= row
    If wkSheet.Cells(i, 1) <> "" Then
        dataArr(Loc, 0) = i     '保存数据在sheet中对应的行号
        '全部数据载入
        For j = 1 To col
            dataArr(Loc, j) = wkSheet.Cells(i, j)
        Next
        '指针移位
        Loc = Loc + 1
    End If
    i = i + 1
Loop
ReDim datas(Loc - 1, col)
'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To col
        datas(i, j) = dataArr(i, j)
    Next
Next
End Sub



Sub 加载竞彩网比分数据(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称

'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId
Dim dataArr
Dim currDate


Dim colCnt As Long   '要加载的数据列总数

Dim leagueData()
Dim league

'取对应关系数据
Call loadLeagueData(leagueData)

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row


colCnt = 36    '此数值为最后一列的列数减1
ReDim dataArr(cnt - 1, colCnt)
'将数据取到内存数组中
Loc = 1

For i = 1 To cnt - 1
    If wkSheet.Cells(i, 1) <> "" Then
        '判断是否要求选择的联赛比赛，通过字典进行判断，加快速度
        league = wkSheet.Cells(i + 1, 4) '赛事
        league = UniformLeague(leagueData, league, 4)
        If leagueDict.exists(league) Then
            dataArr(Loc, 0) = wkSheet.Cells(i + 1, 1) 'ID
            dataArr(Loc, 1) = CDate(wkSheet.Cells(i + 1, 2)) '日期
            dataArr(Loc, 2) = wkSheet.Cells(i + 1, 3)  '编号
            dataArr(Loc, 3) = league  '赛事
            dataArr(Loc, 4) = Trim(wkSheet.Cells(i + 1, 5))  '主队
            dataArr(Loc, 5) = Trim(wkSheet.Cells(i + 1, 6))  '客队
            
            '循环取后续的数值型数据
            For j = 6 To colCnt
                dataArr(Loc, j) = wkSheet.Cells(i + 1, j + 1)
            Next
            
            Loc = Loc + 1
        End If
    End If
Next

ReDim datas(Loc - 1, colCnt)

'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To colCnt
        datas(i, j) = dataArr(i, j)
    Next
Next

Set wkSheet = Nothing
End Sub


Sub 加载竞彩网总进球数据(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称

'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId
Dim dataArr
Dim currDate


Dim colCnt As Long   '要加载的数据列总数

Dim leagueData()
Dim league

'取对应关系数据
Call loadLeagueData(leagueData)

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row


colCnt = 13    '此数值为最后一列的列数减1
ReDim dataArr(cnt - 1, colCnt)
'将数据取到内存数组中
Loc = 1

For i = 1 To cnt - 1
    If wkSheet.Cells(i, 1) <> "" Then
        '判断是否要求选择的联赛比赛，通过字典进行判断，加快速度
        league = wkSheet.Cells(i + 1, 4) '赛事
        league = UniformLeague(leagueData, league, 4)
        If leagueDict.exists(league) Then
            dataArr(Loc, 0) = wkSheet.Cells(i + 1, 1) 'ID
            dataArr(Loc, 1) = CDate(wkSheet.Cells(i + 1, 2)) '日期
            dataArr(Loc, 2) = wkSheet.Cells(i + 1, 3)  '编号
            dataArr(Loc, 3) = league  '赛事
            dataArr(Loc, 4) = Trim(wkSheet.Cells(i + 1, 5))  '主队
            dataArr(Loc, 5) = Trim(wkSheet.Cells(i + 1, 6))  '客队
            
            '循环取后续的数值型数据
            For j = 6 To colCnt
                dataArr(Loc, j) = wkSheet.Cells(i + 1, j + 1)
            Next
            
            Loc = Loc + 1
        End If
    End If
Next

ReDim datas(Loc - 1, colCnt)

'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To colCnt
        datas(i, j) = dataArr(i, j)
    Next
Next

Set wkSheet = Nothing
End Sub


Sub 加载竞彩网半全场胜平负数据(datas, sheetName As String)
'------------------------------------------------------------------------
'dataArr：数据存储的数组
'sheetName：读取的SHEET页名称

'------------------------------------------------------------------------
Dim wkSheet As Worksheet
Dim cnt As Long
Dim Loc As Long
Dim i As Long
Dim j As Long
Dim itemId
Dim dataArr
Dim currDate


Dim colCnt As Long   '要加载的数据列总数

Dim leagueData()
Dim league

'取对应关系数据
Call loadLeagueData(leagueData)

Set wkSheet = ActiveWorkbook.Sheets(sheetName)
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row


colCnt = 14    '此数值为最后一列的列数减1
ReDim dataArr(cnt - 1, colCnt)
'将数据取到内存数组中
Loc = 1

For i = 1 To cnt - 1
    If wkSheet.Cells(i, 1) <> "" Then
        '判断是否要求选择的联赛比赛，通过字典进行判断，加快速度
        league = wkSheet.Cells(i + 1, 4) '赛事
        league = UniformLeague(leagueData, league, 4)
        If leagueDict.exists(league) Then
            dataArr(Loc, 0) = wkSheet.Cells(i + 1, 1) 'ID
            dataArr(Loc, 1) = CDate(wkSheet.Cells(i + 1, 2)) '日期
            dataArr(Loc, 2) = wkSheet.Cells(i + 1, 3)  '编号
            dataArr(Loc, 3) = league  '赛事
            dataArr(Loc, 4) = Trim(wkSheet.Cells(i + 1, 5))  '主队
            dataArr(Loc, 5) = Trim(wkSheet.Cells(i + 1, 6))  '客队
            
            '循环取后续的数值型数据
            For j = 6 To colCnt
                dataArr(Loc, j) = wkSheet.Cells(i + 1, j + 1)
            Next
            
            Loc = Loc + 1
        End If
    End If
Next

ReDim datas(Loc - 1, colCnt)

'数据移植至输出数组
For i = 1 To Loc - 1
    For j = 0 To colCnt
        datas(i, j) = dataArr(i, j)
    Next
Next

Set wkSheet = Nothing
End Sub



