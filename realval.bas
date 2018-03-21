Attribute VB_Name = "realval"
Option Explicit

Sub 实力值计算()
'******************************************************************************************
'
'实力值数据更新
'
'******************************************************************************************

Dim wkWorkbook As Workbook
Dim wkSheet As Worksheet
Dim dataSheet As Worksheet

Dim realSheet As Worksheet
Dim vsSheet As Worksheet
Dim leagueSheet As Worksheet

Dim data()
Dim realData()   '实力值数据
Dim vsData()     '比赛数据


Dim dataConfig()   '配置数据信息

Dim i As Long

'赛事信息
Dim priId As String
Dim secId As String
Dim teamSeason As String
Dim lunci As Integer
Dim teamId As String
Dim slunci As Integer    '每赛季总轮次

Dim lastLunci As Integer   '上一轮次
Dim lastSeason As String    '上赛季
Dim season, tloc

'实力值
Dim priRealLy     '上年主队实力
Dim priRealLast   '上轮主队实力
Dim priReal       '主队实力
Dim secRealLy     '上年客队实力
Dim secRealLast   '上轮客队实力
Dim secReal       '客队实力

Dim vsDict As Object    '比赛数据字典    通过数据直接取比赛相关信息
Dim realDict As Object
Dim leagueDict As Object

'参数

Dim LyRealCol   As Integer     '上年实力值数据开始列号
Dim LastRealcol  As Integer      '上轮实力值数据开始列号
Dim RealCol  As Integer      '本轮数据开始列号

Dim resultCol As Long       '比分数据开始列
Dim winloseCol As Long      '比赛结果数据开始列
Dim turnsCol As Long        '轮次数据开始列


Dim row1, Loc

Dim dealRecCount As Long        '综合数据加载到内存的记录条数
Dim dataBgCol As Long

Dim realFileName As String    '实力值存储文件

Dim priPercent           '实力差值计算中主队分值占比
Dim secPercent           '实力差值计算中客队分值占比



Call 初始化字典(Dict, "Param")

dealRecCount = CLng(Dict.Item("DEAL_RECCOUNT"))
dataBgCol = CLng(Dict.Item("DATABGCOL"))
realFileName = Dict.Item("REAL_FILE_NAME")

priPercent = 1.07
secPercent = 1.05

'将综合数据加载入内存

Set wkSheet = ActiveWorkbook.Sheets("综合数据")
Call 综合数据载入内存(data, "综合数据", dealRecCount, dataBgCol)
tloc = UBound(data)

Set dataSheet = ActiveWorkbook.Sheets("综合数据")
Call 初始化一般字典(dataColDict, dataSheet, 4, 0, 1, False)


LyRealCol = dataColDict.Item("STRENGTHLY")
LastRealcol = dataColDict.Item("STRENGTHLAST")
RealCol = dataColDict.Item("STRENGTH")

resultCol = dataColDict.Item("RESULT")
winloseCol = dataColDict.Item("WINLOSE")
turnsCol = dataColDict.Item("TURNS")

'加载实力值相关数据
On Error Resume Next  '遇到错误继续执行下一行
Set wkWorkbook = Application.Workbooks("球队实力.xlsm")
If wkWorkbook Is Nothing Then
    Set wkWorkbook = Application.Workbooks.Open(ActiveWorkbook.path + "\" + realFileName) '球队实力.xlsm")
End If
Set realSheet = wkWorkbook.Sheets("球队实力")

If realSheet Is Nothing Then
    MsgBox ("请打开【球队实力】文件！")
    Exit Sub
End If

Set vsSheet = wkWorkbook.Sheets("赛程积分")
Set leagueSheet = wkWorkbook.Sheets("LeagueConfig")


Call 数据载入通用程序(vsData, vsSheet, 0, 2)
Call 数据载入通用程序(realData, realSheet, dealRecCount, 2)

Call 初始化球队实力字典(realDict, realSheet)

Call 从数组构建字典(vsDict, vsData, 3, 0, 1)
Call 初始化一般字典(leagueDict, leagueSheet, 1, 3, 2)

For i = 1 To UBound(data)
    row1 = data(i, 0)  '取出行号
    
    Loc = vsDict.Item(data(i, 8))     '找到实力值表中对应的位置
    '取出主队Id、客队Id、赛季、轮次
    If Loc > 0 Then
        teamId = vsData(Loc, 4)    '联赛Id
        priId = vsData(Loc, 7)     '主队Id
        secId = vsData(Loc, 8)     '客队Id
        teamSeason = vsData(Loc, 1)  '赛季
        lunci = vsData(Loc, 2)       '轮次
        
        '求上一赛季
        season = Split(teamSeason, "-")
        lastSeason = CStr(CInt(season(0)) - 1) & "-" & CStr(CInt(season(1)) - 1)
        
        slunci = leagueDict.Item(CDbl(teamId))
        lastLunci = lunci - 1
        
        '取出实力值,根据字典（teamId & priId & teamSeason & lunci)
        priRealLy = realDict.Item(teamId & priId & lastSeason & slunci)
        
        priReal = realDict.Item(teamId & priId & teamSeason & lunci)
        
        
        secRealLy = realDict.Item(teamId & secId & lastSeason & slunci)
        secReal = realDict.Item(teamId & secId & teamSeason & lunci)
        If lunci > 1 Then
            priRealLast = realDict.Item(teamId & priId & teamSeason & lastLunci)
            secRealLast = realDict.Item(teamId & secId & teamSeason & lastLunci)
        Else
            priRealLast = 100
            secRealLast = 100
        End If
        '写出excel中
        
        '上年实力值
        wkSheet.Cells(row1, LyRealCol) = priRealLy
        wkSheet.Cells(row1, LyRealCol + 1) = secRealLy
        wkSheet.Cells(row1, LyRealCol + 2) = priRealLy * priPercent - secRealLy * secPercent
        
        '上轮实力值
        wkSheet.Cells(row1, LastRealcol) = priRealLast
        wkSheet.Cells(row1, LastRealcol + 1) = secRealLast
        wkSheet.Cells(row1, LastRealcol + 2) = priRealLast * priPercent - secRealLast * secPercent
        
        '本轮实力值
        If InStr(vsData(Loc, 9), "-") > 0 Then
            wkSheet.Cells(row1, RealCol) = priReal
            wkSheet.Cells(row1, RealCol + 1) = secReal
            wkSheet.Cells(row1, RealCol + 2) = priReal * priPercent - secReal * secPercent
            '全场得分
            wkSheet.Cells(row1, resultCol) = vsData(Loc, 9)
            wkSheet.Cells(row1, winloseCol) = 比赛结果(vsData(Loc, 9))
            '半场得分
            wkSheet.Cells(row1, resultCol + 1) = vsData(Loc, 10)
            wkSheet.Cells(row1, winloseCol + 1) = 比赛结果(vsData(Loc, 10))
        End If
        
        '轮次数据
        wkSheet.Cells(row1, turnsCol) = lunci
        
    End If
Next i


Set realSheet = Nothing
Set vsSheet = Nothing
Set leagueSheet = Nothing

Set vsDict = Nothing
Set realDict = Nothing
Set leagueDict = Nothing
Set wkWorkbook = Nothing

MsgBox ("实力值传送完毕！")

End Sub


Sub 初始化球队实力字典(tempDict As Object, paraSheet As Worksheet)
Dim itemId, itemVal
Dim dcnt
Dim cnt, i

On Error Resume Next  '遇到错误继续执行下一行
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If


'初始化参数

cnt = paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row
For i = 2 To cnt
    If paraSheet.Cells(i, 1) <> "" And paraSheet.Cells(i, 3) <> "" And paraSheet.Cells(i, 5) <> "" And paraSheet.Cells(i, 5) <> "" Then
        itemId = paraSheet.Cells(i, 1) & paraSheet.Cells(i, 3) & paraSheet.Cells(i, 5) & paraSheet.Cells(i, 6)
        itemVal = paraSheet.Cells(i, 7)
        
        If tempDict.exists(itemId) Then
            tempDict.Item(itemId) = itemVal
        Else
            tempDict.Add itemId, itemVal
        End If
    End If
Next
End Sub
