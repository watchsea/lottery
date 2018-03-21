Attribute VB_Name = "userdef"
Option Explicit
Function 模式四(i As Double, j As Double, k As Double)
Dim str As String
str = ""
    If i > j Then
        If j > k Then
            If j > 0 Then    'i>j>k  and i,j>0;k<0
                str = "±+-"
            Else            'i>j>k  and i>0;j,k<0
                str = "+-="
            End If
        Else   'j<=k
            If i > k Then    'i>k>j
                If k > 0 Then       'i>k>j,  i,k>0; j<0
                    str = "±-+"
                Else                'i>k>j,  i>0,k,j<0
                    str = "+=-"
                End If
            Else             'k>i>j
                If i > 0 Then       'k>i>j,k,i>0,j<0
                    str = "+-±"
                Else               'k>i>j,k>0,i,j<0
                    str = "-=+"
                End If
            End If
        End If
    ElseIf i < j Then 'i<j'
        If j < k Then     'i<j<k
            If j > 0 Then    'i<j<k,j,k>0,i<0
                str = "-+±"
            Else             'i<j<k,k>0,i,j<0
                str = "=-+"
            End If
        Else 'j>k
            If i < k Then   'j>k>i
                If k > 0 Then     'j>k>i,j,k>0,i<0
                    str = "-±+"
                Else              'j>k>i,j>0,k,i<0
                    str = "=+-"
                End If
            Else           'j>i>k
                If i > 0 Then    'j>i>k,  j,i>0,k<0
                    str = "+±-"
                Else             'j>i>k,  j>0,i,k<0
                    str = "-+="
                End If
            End If
        End If
    Else   'i=j
        
    End If

模式四 = str
End Function



Function 标识(i As Double, j As Double, k As Double)
'i :代表胜，j:代表平，k:代表负
Dim str As String
str = ""
    If i < j And j < k Then str = "A"
    If i > j And j > k Then str = "-A"
    If j > i And j > k And i < k Then str = "D"
    If j > i And j > k And i > k Then str = "-D"
    If j > i And j > k And i = k Then str = "-E"
    If j < i And j < k And i < k Then str = "B"
    If j < i And j < k And i > k Then str = "-B"
    If j < i And j < k And i = k Then str = "E"
    If i = j And i < k Then str = "G"
    If i = j And i > k Then str = "-C"
    If j = k And j < i Then str = "-G"
    If j = k And j > i Then str = "C"
    If i = j And j = k And i <> 0 Then str = "F"
    If i = j And j = k And i = 0 Then str = ""
标识 = str
End Function


Function 比较(dataSheet, rowNo, colNo, offset, compareType As String)
'------------------------------------------------
'dataSheet 对应的数据页
'rowNo,colNo,待操作的数据单元
'offset :列偏移量
'compareType:按什么方式进行比较，"D"，对小于0的数据进行“升序”排列；
'                                "A",对于大于0的数据进行“降序”排列
Dim str As String
Dim data(0, 2)
Dim index(0, 2)
Dim colDesc
Dim sortType  As String
Dim tempData

Dim i, j, k

'当前数据为空，则比较栏填入空
If dataSheet.Cells(rowNo, colNo - offset) = "" And dataSheet.Cells(rowNo, colNo + 1 - offset) = "" And dataSheet.Cells(rowNo, colNo + 2 - offset) = "" Then
    dataSheet.Cells(rowNo, colNo).Value = ""
    Exit Function
End If

'如果上下行数据相等
If dataSheet.Cells(rowNo, colNo - offset) = dataSheet.Cells(rowNo - 1, colNo - offset) And dataSheet.Cells(rowNo, colNo + 1 - offset) = dataSheet.Cells(rowNo - 1, colNo + 1 - offset) And dataSheet.Cells(rowNo, colNo + 2 - offset) = dataSheet.Cells(rowNo - 1, colNo + 2 - offset) Then
    dataSheet.Cells(rowNo, colNo).Value = ""
    Exit Function
End If

i = dataSheet.Cells(rowNo, colNo - offset) - dataSheet.Cells(rowNo - 1, colNo - offset)
j = dataSheet.Cells(rowNo, colNo + 1 - offset) - dataSheet.Cells(rowNo - 1, colNo + 1 - offset)
k = dataSheet.Cells(rowNo, colNo + 2 - offset) - dataSheet.Cells(rowNo - 1, colNo + 2 - offset)



colDesc = Split("3,1,0", ",")
If compareType = "D" Then '取减少的数据
    If i > 0 Then data(0, 0) = 0 Else data(0, 0) = i
    If j > 0 Then data(0, 1) = 0 Else data(0, 1) = j
    If k > 0 Then data(0, 2) = 0 Else data(0, 2) = k
    sortType = "A"
Else     '默认取增加的数据
    If i < 0 Then data(0, 0) = 0 Else data(0, 0) = i
    If j < 0 Then data(0, 1) = 0 Else data(0, 1) = j
    If k < 0 Then data(0, 2) = 0 Else data(0, 2) = k
    sortType = "D"
End If

Call SortCompareData(data, index, sortType)
tempData = 生成模式符号(data, index, colDesc, 4)
dataSheet.Cells(rowNo, colNo).NumberFormatLocal = "@"
dataSheet.Cells(rowNo, colNo).Value = tempData

End Function



Function 横向比较(dataSheet, rowNo, colNo, offset, compareType As String, Optional lbl As Integer = 1)
'------------------------------------------------
'dataSheet 对应的数据页
'rowNo,colNo,待操作的数据单元
'offset :列偏移量
'compareType:按什么方式进行比较，"D"，对小于0的数据进行“升序”排列；
'                                "A",对于大于0的数据进行“降序”排列
'lbl:比较标识说明： 1：默认值，横向与某项数据进行比较，offset为两个数据项的偏移量
'                   2：某项数据内部进行比较，offset为相同减项与数据的偏移量
Dim str As String
Dim data(0, 2)
Dim index(0, 2)
Dim colDesc
Dim sortType  As String
Dim tempData

Dim i, j, k

If lbl = 2 Then    '数据项内进行比较
    i = dataSheet.Cells(rowNo, colNo - offset - 3) - dataSheet.Cells(rowNo, colNo - offset)
    j = dataSheet.Cells(rowNo, colNo - offset - 2) - dataSheet.Cells(rowNo, colNo - offset)
    k = dataSheet.Cells(rowNo, colNo - offset - 1) - dataSheet.Cells(rowNo, colNo - offset)
Else     '两数据项进行比较
    i = dataSheet.Cells(rowNo, colNo - 3) - dataSheet.Cells(rowNo, colNo - offset - 3)
    j = dataSheet.Cells(rowNo, colNo - 2) - dataSheet.Cells(rowNo, colNo - offset - 2)
    k = dataSheet.Cells(rowNo, colNo - 1) - dataSheet.Cells(rowNo, colNo - offset - 1)
End If

colDesc = Split("3,1,0", ",")
If compareType = "D" Then '取减少的数据
    If i > 0 Then data(0, 0) = 0 Else data(0, 0) = i
    If j > 0 Then data(0, 1) = 0 Else data(0, 1) = j
    If k > 0 Then data(0, 2) = 0 Else data(0, 2) = k
    sortType = "A"
Else     '默认取增加的数据
    If i < 0 Then data(0, 0) = 0 Else data(0, 0) = i
    If j < 0 Then data(0, 1) = 0 Else data(0, 1) = j
    If k < 0 Then data(0, 2) = 0 Else data(0, 2) = k
    sortType = "D"
End If

Call SortCompareData(data, index, sortType)
tempData = 生成模式符号(data, index, colDesc, 4)
横向比较 = tempData
End Function


Function 固定值比较(i1, j1, k1, fixValue, compareType As String)
'------------------------------------------------
'dataSheet 对应的数据页
'rowNo,colNo,待操作的数据单元
'offset :列偏移量
'compareType:按什么方式进行比较，"D"，对小于0的数据进行“升序”排列；
'                                "A",对于大于0的数据进行“降序”排列
'lbl:比较标识说明： 1：默认值，横向与某项数据进行比较，offset为两个数据项的偏移量
'                   2：某项数据内部进行比较，offset为相同减项与数据的偏移量
Dim str As String
Dim data(0, 2)
Dim index(0, 2)
Dim colDesc
Dim sortType  As String
Dim tempData
Dim i, j, k

i = i1 - fixValue
j = j1 - fixValue
k = k1 - fixValue

colDesc = Split("3,1,0", ",")
If compareType = "D" Then '取减少的数据
    If i > 0 Then data(0, 0) = 0 Else data(0, 0) = i
    If j > 0 Then data(0, 1) = 0 Else data(0, 1) = j
    If k > 0 Then data(0, 2) = 0 Else data(0, 2) = k
    sortType = "A"
Else     '默认取增加的数据
    If i < 0 Then data(0, 0) = 0 Else data(0, 0) = i
    If j < 0 Then data(0, 1) = 0 Else data(0, 1) = j
    If k < 0 Then data(0, 2) = 0 Else data(0, 2) = k
    sortType = "D"
End If

Call SortCompareData(data, index, sortType)
tempData = 生成模式符号(data, index, colDesc, 4)
固定值比较 = tempData
End Function



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

Set x1 = ActiveWorkbook.Sheets("01赛事")

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


