Attribute VB_Name = "userdefine"
Option Explicit

Function UniformLeague(leagueData, league, colNo)
'统一联赛名称，以便匹配数据
'leagueData ：联赛对应关系数据
'league：联赛
'netName：联赛数据中对应的列号
Dim i, j
For i = 1 To UBound(leagueData, 1)   '行
    If league = leagueData(i, colNo) Then
        Exit For
    End If
Next

If i <= UBound(leagueData, 1) Then
    UniformLeague = leagueData(i, 1)
Else
    UniformLeague = league
End If

End Function

Sub loadLeagueData(leagueData())
'取各网站联赛名称对应关系数据

Dim x1 As Worksheet
Dim colNo As Integer
Dim rowNo As Integer
Dim i, j
Dim cnt

Set x1 = ThisWorkbook.Sheets("LeagueConfig")

rowNo = x1.UsedRange.Rows(x1.UsedRange.Rows.Count).row
colNo = x1.UsedRange.Columns(x1.UsedRange.Columns.Count).Column

ReDim leagueData(rowNo - 1, colNo)
cnt = 0
For i = 2 To rowNo
    If x1.Cells(i, 1) <> "" Then
        cnt = cnt + 1
        For j = 1 To colNo
            leagueData(cnt, j) = x1.Cells(i, j)
        Next
    End If

Next
Set x1 = Nothing
End Sub



Sub SortCompareData(iSortData, sortIndex, Optional sortType As String = "A")
'对几组数据进行排序，
'sortData 是待排序的二维数组
'sortIndex 是保存排序后的索引
'rowOrCol: 行列排序说明：R:对行进行排序，C：对列进行排序
'sortType 是排序类型：A：升序，D：降序
Dim i, j, k
Dim rowLen, colLen
Dim tempData
Dim tempIndex
Dim sortData1()

sortData1 = iSortData

rowLen = UBound(sortData1, 1)
colLen = UBound(sortData1, 2)

For i = 0 To rowLen
     '以下对每一行的数据进行排序，排序结果将数据序号保存在sortIndex对应的行中

    For j = 0 To colLen
        tempData = sortData1(i, j)
        tempIndex = j

        For k = 0 To colLen

            If sortType = "D" Then   '降序
                If sortData1(i, k) > tempData Then
                    tempData = sortData1(i, k)
                    tempIndex = k
                End If
            Else     '默认升序
                If sortData1(i, k) < tempData Then
                    tempData = sortData1(i, k)
                    tempIndex = k
                End If
            End If

        Next
        sortIndex(i, j) = tempIndex
        If sortType = "D" Then
            sortData1(i, tempIndex) = -1
        Else
            sortData1(i, tempIndex) = 1
        End If
    Next
Next
End Sub

Function 比赛结果(result, Optional separator As String = "-")
'根据比赛分数，计算比赛结果
'result:用于输入比分

Dim r1
Dim a1 As Integer
Dim a2 As Integer
Dim str As String
r1 = Split(result, separator)
If UBound(r1) <> 1 Then
    str = ""
Else
    a1 = CInt(r1(0))
    a2 = CInt(r1(1))
    If a1 > a2 Then
        str = "3"
    ElseIf a1 < a2 Then
        str = "0"
    Else
        str = "1"
    End If
End If
比赛结果 = str
End Function


