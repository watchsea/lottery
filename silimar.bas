Attribute VB_Name = "silimar"
Option Explicit

Sub 相同赔率比较(ByRef control As Office.IRibbonControl)
'此模块主要是处理相同赔率的数据比较
'
'
Dim outWorkbook As Workbook
Dim wkWorkbook As Workbook
Dim wkSheet As Worksheet
Dim outSheet As Worksheet
Dim outSheet2 As Worksheet

Dim path As String    '路径
Dim outSheetName As String


Dim outDict As Object     '输出Sheet的主键字典
Dim outDict2 As Object    '输出Sheet2的主键字典
Dim srcdata()
Dim tgtdata()
Dim tgtdata2()

Dim id As String     '对阵ID
Dim res2 As String    '比较数据2的比赛结果


Dim row1 As Long
Dim col1 As Long

Dim row2 As Long
Dim col2 As Long

Dim row3 As Long
Dim col3 As Long



Dim cnt As Long    '序号
Dim row As Long    '指针
Dim tmpCol As Long  '临时指针

'胜平负数据
Dim win1     '胜
Dim eq1      '平
Dim lose1    '负
Dim win2
Dim eq2
Dim lose2
Dim result As String   '结果
Dim i, j


Dim updateRow As Long   '待插入的位置
Dim updateRow2 As Long   '表二待插入的位置
Dim updateFlag As Boolean   '是否需要修改数据


Dim dataWcol As Integer   '威廉希尔数据开始列号
Dim dataBcol As Integer     'Bet365数据开始列号
Dim dataMcol As Integer     '澳门彩票数据开始列号
Dim dataLcol As Integer     '立博(英国)数据开始列号
Dim dataEcol As Integer     '易胜博数据开始列号
Dim DataJcol As Integer     '竞彩网数据开始列号

Dim lose2Col As Integer     '赔2 数据开始列号
Dim OKBF1col As Integer     'OKBF1数据开始列号
Dim OKBF2col As Integer     'OKBF2数据开始列号
Dim off1 As Integer         '输出的偏移量
Dim srcBgCol As Integer     '原文件对应项目开始列
Dim dealRecCount As Long    '处理数据条数

Dim cmpFileName As String    '相同比较的文件名


Call 初始化字典(Dict, "Param")



dealRecCount = CLng(Dict.Item("COMP_CNT"))
cmpFileName = Dict.Item("CMP_FILE_NAME")



Set wkWorkbook = ActiveWorkbook
Set wkSheet = wkWorkbook.Sheets("综合数据")

Call 初始化一般字典(dataColDict, wkSheet, 4, 0, 1, False)

dataWcol = dataColDict.Item("DATAW")
dataBcol = dataColDict.Item("DATAB")
dataMcol = dataColDict.Item("DATAM")
dataLcol = dataColDict.Item("DATAL")
dataEcol = dataColDict.Item("DATAE")
DataJcol = dataColDict.Item("DATAJ")


lose2Col = dataColDict.Item("LOSE2")
OKBF1col = dataColDict.Item("OKBF1")
OKBF2col = dataColDict.Item("OKBF2")


row1 = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
col1 = wkSheet.UsedRange.Columns(wkSheet.UsedRange.Columns.Count).Column

path = ActiveWorkbook.path

outSheetName = path + "\" + cmpFileName '相同赔率比较.xlsx"
Set outWorkbook = Workbooks.Open(outSheetName)
Set outSheet = outWorkbook.Sheets("sheet1")
row2 = outSheet.UsedRange.Rows(outSheet.UsedRange.Rows.Count).row
col2 = outSheet.UsedRange.Columns(outSheet.UsedRange.Columns.Count).Column


Set outSheet2 = outWorkbook.Sheets("sheet2")
row3 = outSheet2.UsedRange.Rows(outSheet2.UsedRange.Rows.Count).row
col3 = outSheet2.UsedRange.Columns(outSheet2.UsedRange.Columns.Count).Column


'从【综合数据】表中取出数据到内存中
Call 数据载入通用程序(srcdata, wkSheet, 0, 4)

'把相同赔率比较表中对阵ID写入字典，以便检查
Call 初始化一般字典(outDict, outSheet, 3, 0, 5)

'把相同赔率比较表2中对阵ID写入字典，以便确定数据更新位置
Call 初始化一般字典(outDict2, outSheet2, 3, 0, 5)

ReDim tgtdata(col2)
ReDim tgtdata2(col3)
If row2 = 4 Then
    cnt = 0
Else
    cnt = outSheet.Cells(row2 - 1, 1) '取得最后的序号
End If

row = row2
For i = UBound(srcdata) - 3 * dealRecCount To UBound(srcdata)
    If srcdata(i, 6) = "初始值" Then
        off1 = 4
        id = srcdata(i, 9)    '对阵Id
        If outDict2.Item(id) <> "" Then res2 = outSheet2.Cells(outDict2.Item(id), 4) '比较结果2中的赛事结果
        If Not (outDict.exists(id)) Then   '不存在，则添加
            updateFlag = True
            cnt = cnt + 1
            row2 = row2 + 1
            row = row2
            updateRow = cnt
        ElseIf srcdata(i, 11) = "" Or (srcdata(i, 11) <> "" And res2 = "") Then '如果原始数据的比赛结果为空，或者原始数据的比赛结果已出来，而比较数据还未更新，则需时时更新
            updateFlag = True
            row = outDict.Item(id)    '找到要修改的行号
            updateRow = outSheet.Cells(row, 1)    '取得全行号
        End If
    Else
        updateFlag = False
    End If
    If updateFlag Then
        '添加基本信息
        tgtdata(1) = updateRow
        tgtdata(2) = srcdata(i, 5)   '对阵名称
        tgtdata(3) = id
        
        tgtdata2(1) = updateRow
        tgtdata2(2) = srcdata(i, 5) '对阵名称
        tgtdata2(3) = id
        tgtdata2(4) = srcdata(i, 11)  '比赛结果
        '预置初始值
        For j = 4 To col2
            tgtdata(j) = ""
        Next
        '预置比较表二的初始值
        For j = 5 To col3
            tgtdata2(j) = ""
        Next
        For j = 1 To i - 1  '统计相同赔率的信息
            If srcdata(j, 6) = "初始值" Then
                '威廉希尔
                tmpCol = 4
                srcBgCol = dataWcol
                If srcdata(j, srcBgCol) = srcdata(i, srcBgCol) And srcdata(j, srcBgCol + 1) = srcdata(i, srcBgCol + 1) And srcdata(j, srcBgCol + 2) = srcdata(i, srcBgCol + 2) Then
                    Call 求相同赔率比较值(tgtdata, tmpCol, 0, off1, srcdata, j, lose2Col, True)  '赔2
                    Call 求相同赔率比较值(tgtdata, tmpCol, 1, off1, srcdata, j, OKBF1col, False)  'OKBF1
                    Call 求相同赔率比较值(tgtdata, tmpCol, 2, off1, srcdata, j, OKBF2col, False)   'OKBF2
                    
                    
                    '处理下面的细分类
                    Call 相同赔率比较细分(tgtdata2, srcdata, CLng(i), CLng(j), Dict, "W")
                    
                End If
                
                'Bet365
                tmpCol = 16
                srcBgCol = dataBcol
                If srcdata(j, srcBgCol) = srcdata(i, srcBgCol) And srcdata(j, srcBgCol + 1) = srcdata(i, srcBgCol + 1) And srcdata(j, srcBgCol + 2) = srcdata(i, srcBgCol + 2) Then
                     Call 求相同赔率比较值(tgtdata, tmpCol, 0, off1, srcdata, j, lose2Col, True)  '赔2
                    Call 求相同赔率比较值(tgtdata, tmpCol, 1, off1, srcdata, j, OKBF1col, False)  'OKBF1
                    Call 求相同赔率比较值(tgtdata, tmpCol, 2, off1, srcdata, j, OKBF2col, False)   'OKBF2
                    
                     '处理下面的细分类
                    Call 相同赔率比较细分(tgtdata2, srcdata, CLng(i), CLng(j), Dict, "B")
                End If
                '澳门
                tmpCol = 28
                srcBgCol = dataMcol
                If srcdata(j, srcBgCol) = srcdata(i, srcBgCol) And srcdata(j, srcBgCol + 1) = srcdata(i, srcBgCol + 1) And srcdata(j, srcBgCol + 2) = srcdata(i, srcBgCol + 2) Then
                     Call 求相同赔率比较值(tgtdata, tmpCol, 0, off1, srcdata, j, lose2Col, True)  '赔2
                    Call 求相同赔率比较值(tgtdata, tmpCol, 1, off1, srcdata, j, OKBF1col, False)  'OKBF1
                    Call 求相同赔率比较值(tgtdata, tmpCol, 2, off1, srcdata, j, OKBF2col, False)   'OKBF2
                    
                     '处理下面的细分类
                    Call 相同赔率比较细分(tgtdata2, srcdata, CLng(i), CLng(j), Dict, "M")
                    
                End If
                '中国彩票网
                tmpCol = 40
                srcBgCol = DataJcol
                If srcdata(j, srcBgCol) = srcdata(i, srcBgCol) And srcdata(j, srcBgCol + 1) = srcdata(i, srcBgCol + 1) And srcdata(j, srcBgCol + 2) = srcdata(i, srcBgCol + 2) Then
                     Call 求相同赔率比较值(tgtdata, tmpCol, 0, off1, srcdata, j, lose2Col, True)  '赔2
                    Call 求相同赔率比较值(tgtdata, tmpCol, 1, off1, srcdata, j, OKBF1col, False)  'OKBF1
                    Call 求相同赔率比较值(tgtdata, tmpCol, 2, off1, srcdata, j, OKBF2col, False)   'OKBF2
                
                End If
                '立博
                tmpCol = 52
                srcBgCol = dataLcol
                If srcdata(j, srcBgCol) = srcdata(i, srcBgCol) And srcdata(j, srcBgCol + 1) = srcdata(i, srcBgCol + 1) And srcdata(j, srcBgCol + 2) = srcdata(i, srcBgCol + 2) Then
                     Call 求相同赔率比较值(tgtdata, tmpCol, 0, off1, srcdata, j, lose2Col, True)  '赔2
                    Call 求相同赔率比较值(tgtdata, tmpCol, 1, off1, srcdata, j, OKBF1col, False)  'OKBF1
                    Call 求相同赔率比较值(tgtdata, tmpCol, 2, off1, srcdata, j, OKBF2col, False)   'OKBF2
                    
                     '处理下面的细分类
                    Call 相同赔率比较细分(tgtdata2, srcdata, CLng(i), CLng(j), Dict, "L")
                End If
                '易胜博
                tmpCol = 64
                srcBgCol = dataEcol
                If srcdata(j, srcBgCol) = srcdata(i, srcBgCol) And srcdata(j, srcBgCol + 1) = srcdata(i, srcBgCol + 1) And srcdata(j, srcBgCol + 2) = srcdata(i, srcBgCol + 2) Then
                     Call 求相同赔率比较值(tgtdata, tmpCol, 0, off1, srcdata, j, lose2Col, True)  '赔2
                    Call 求相同赔率比较值(tgtdata, tmpCol, 1, off1, srcdata, j, OKBF1col, False)  'OKBF1
                    Call 求相同赔率比较值(tgtdata, tmpCol, 2, off1, srcdata, j, OKBF2col, False)   'OKBF2
                    
                    
                     '处理下面的细分类
                    Call 相同赔率比较细分(tgtdata2, srcdata, CLng(i), CLng(j), Dict, "E")
                End If
            End If    '如果是初始值判断结束
        Next   '内循环完毕
        
        
        '计算结果
        '---------------------赔2---------------------------------
        tmpCol = 4
        off1 = 12
        
        For j = 0 To 5
            result = ""
            If srcdata(i, lose2Col) > tgtdata(tmpCol + j * off1) And tgtdata(tmpCol + j * off1) <> 0 Then ' 胜
                result = result + "3"
            End If
            If srcdata(i, lose2Col + 1) > tgtdata(tmpCol + j * off1 + 1) And tgtdata(tmpCol + j * off1 + 1) <> 0 Then ' 平
                result = result + "1"
            End If
            If srcdata(i, lose2Col + 2) > tgtdata(tmpCol + j * off1 + 2) And tgtdata(tmpCol + j * off1 + 2) <> 0 Then ' 负
                result = result + "0"
            End If
            tgtdata(tmpCol + j * off1 + 3) = result
        Next
                
        '------------------OKBF1------------------------------------
        tmpCol = 8
        off1 = 12
        
        For j = 0 To 5
            result = ""
            If srcdata(i, OKBF1col) < tgtdata(tmpCol + j * off1) And srcdata(i, OKBF1col) <> 0 Then ' 胜
                result = result + "3"
            End If
            If srcdata(i, OKBF1col + 1) < tgtdata(tmpCol + j * off1 + 1) And srcdata(i, OKBF1col + 1) <> 0 Then ' 平
                result = result + "1"
            End If
            If srcdata(i, OKBF1col + 2) < tgtdata(tmpCol + j * off1 + 2) And srcdata(i, OKBF1col + 2) <> 0 Then ' 负
                result = result + "0"
            End If
            tgtdata(tmpCol + j * off1 + 3) = result
        Next
                
        '------------------OKBF2------------------------------------
        tmpCol = 12
        off1 = 12
        For j = 0 To 5
            result = ""
            If srcdata(i, OKBF2col) < tgtdata(tmpCol + j * off1) And srcdata(i, OKBF2col) <> 0 Then ' 胜
                result = result + "3"
            End If
            If srcdata(i, OKBF2col + 1) < tgtdata(tmpCol + j * off1 + 1) And srcdata(i, OKBF2col + 1) <> 0 Then ' 平
                result = result + "1"
            End If
            If srcdata(i, OKBF2col + 2) < tgtdata(tmpCol + j * off1 + 2) And srcdata(i, OKBF2col + 2) <> 0 Then ' 负
                result = result + "0"
            End If
            tgtdata(tmpCol + j * off1 + 3) = result
        Next
        
        
        '表2的赔2数据的计算
        tmpCol = 5
        off1 = 8
        
        For j = 0 To 24   '五个大节，每个大节5个小节，共25节
            result = ""
            If srcdata(i + 2, lose2Col) > tgtdata2(tmpCol + j * off1) And tgtdata2(tmpCol + j * off1) <> 0 Then ' 胜
                result = result + "3"
            End If
            If srcdata(i + 2, lose2Col + 1) > tgtdata2(tmpCol + j * off1 + 1) And tgtdata2(tmpCol + j * off1 + 1) <> 0 Then ' 平
                result = result + "1"
            End If
            If srcdata(i + 2, lose2Col + 2) > tgtdata2(tmpCol + j * off1 + 2) And tgtdata2(tmpCol + j * off1 + 2) <> 0 Then ' 负
                result = result + "0"
            End If
            tgtdata2(tmpCol + j * off1 + 3) = result
        Next
        
        '写入EXCEL
        For j = 1 To col2
            outSheet.Cells(row, j) = tgtdata(j)
        Next
        '默认两个表的指针是同步的
        For j = 1 To col3
            outSheet2.Cells(row, j) = tgtdata2(j)
        Next
        
    End If
Next

outWorkbook.Save

Set outDict = Nothing

Set outSheet = Nothing
Set outWorkbook = Nothing

Set wkWorkbook = Nothing
Set wkWorkbook = Nothing
MsgBox ("输出相同赔率数据结束!")

End Sub



Sub 求相同赔率比较值(tgtdata(), tmpCol, multi, off1, srcdata(), j, srcCol, great As Boolean)
Dim tgtVal As Double
Dim srcVal As Double
Dim cond1 As Boolean
Dim i As Integer
Dim k As Integer


For i = 0 To 2
      tgtVal = 0
    srcVal = 0
    If tgtdata(tmpCol + multi * off1 + i) <> "" Then tgtVal = CDbl(tgtdata(tmpCol + multi * off1 + i))
    If srcdata(j, 11) <> "" Then
        If CStr(srcdata(j, 11)) = "3" Then
            k = 0
        ElseIf CStr(srcdata(j, 11)) = "1" Then
            k = 1
        ElseIf CStr(srcdata(j, 11)) = "0" Then
            k = 2
        Else
            k = -1
        End If
    End If
        
        
    If k = i Then
        If IsNull(srcdata(j, srcCol + i)) And srcdata(j, srcCol + i) = "" Then srcVal = 0 Else srcVal = CDbl(srcdata(j, srcCol + i))
        
        If great = True Then
            If tgtVal > srcVal Then cond1 = True Else cond1 = False
        ElseIf tgtVal < srcVal Then cond1 = True
        Else: cond1 = False
        End If
        
        If cond1 And srcVal > 0 Then
            tgtdata(tmpCol + multi * off1 + i) = srcVal
        ElseIf tgtdata(tmpCol + multi * off1 + i) = "" And Not IsNull(srcdata(j, srcCol + i)) And srcVal > 0 Then
            tgtdata(tmpCol + multi * off1 + i) = srcVal
        End If
    End If
Next
End Sub

Sub 相同赔率比较细分(tgtdata, srcdata, srccol1 As Long, srccol2 As Long, dict1 As Object, class As String)
'------------------------------------------------------------------------------------------
'tgtdata:目标数据
'srcdata:源数据
'srccol1: 源数据基准行的行号
'srccol2:源数据比较行的行号
'dict1:参数字典
'class:   大类说明： 威廉希尔——W,bets365——B，澳门——M，立博——L，易胜博——E
'------------------------------------------------------------------------------------------
Dim schema1Col As Long     '模式1开始列
Dim schema3Col As Long     '模式3开始列
Dim lose2Col As Long       '赔2数据开始列
Dim BF1Col As Long        'BF1凯利开始列
Dim varCol As Long        '方差数据开始列
Dim caliCol As Long       '凯利数据开始列   威廉希尔对应BFW，BETS365 对应BFB，澳门数据对应于 BFM


Dim id As String      '字典ID
Dim itemVal As Long   '字典值


'取出相应的字典项值
schema1Col = CLng(dict1.Item("SCHEMA_COL"))
schema3Col = schema1Col + 2

lose2Col = CLng(dict1.Item("LOSE2_COL"))
BF1Col = CLng(dict1.Item("BF1_COL"))
varCol = CLng(dict1.Item("VAR_COL"))


'处理模式1下数据
If srcdata(srccol1, schema1Col) <> "" And srcdata(srccol1, schema1Col) = srcdata(srccol2, schema1Col) Then
    Call 比较细分按模式(tgtdata, srcdata, 5, 0, 8, dict1, srccol1, srccol2, class)
End If
'处理模式3下数据
If srcdata(srccol1, schema3Col) <> "" And srcdata(srccol1, schema3Col) = srcdata(srccol2, schema3Col) Then
    Call 比较细分按模式(tgtdata, srcdata, 5, 1, 8, dict1, srccol1, srccol2, class)
End If
'处理BF1凯利+方差下数据,即时值2的值
If srcdata(srccol1 + 2, BF1Col + 4) <> "" And srcdata(srccol1 + 2, BF1Col + 4) = srcdata(srccol2 + 2, BF1Col + 4) And srcdata(srccol1 + 2, varCol + 3) <> "" And srcdata(srccol1 + 2, varCol + 3) = srcdata(srccol2 + 2, varCol + 3) Then
    Call 比较细分按模式(tgtdata, srcdata, 5, 2, 8, dict1, srccol1, srccol2, class)
End If
'处理BF1凯利下数据，即时值2的值
If srcdata(srccol1 + 2, BF1Col + 4) <> "" And srcdata(srccol1 + 2, BF1Col + 4) = srcdata(srccol2 + 2, BF1Col + 4) Then
    Call 比较细分按模式(tgtdata, srcdata, 5, 3, 8, dict1, srccol1, srccol2, class)
End If
'处理方差下数据，即时值2的值
If srcdata(srccol1 + 2, varCol + 3) <> "" And srcdata(srccol1 + 2, varCol + 3) = srcdata(srccol2 + 2, varCol + 3) Then
    Call 比较细分按模式(tgtdata, srcdata, 5, 4, 8, dict1, srccol1, srccol2, class)
End If


End Sub


Sub 比较细分按模式(tgtdata, srcdata, tmpCol As Long, multi2 As Long, off2 As Long, dict1 As Object, row1 As Long, j As Long, class As String, Optional great As Boolean = True)
'-------------------------------------------------------------------------------------------------------
'对各类数据来源进行按模式细分
'tgtdata,目标数据
'srcdata,源数据
'tmpCol,目标数据的计算起始列，默认为5
'multi2,大类中的小类的乘数,
'off2，大类中的小类的偏移
'row1,srcdata的基准行，
'j,srcdata的比较行
'class:   大类说明： 威廉希尔——W,bets365——B，澳门——M，立博——L，易胜博——E
'great: 比较类型 ，true——取最小值，false——取最大值
'-------------------------------------------------------------------------------------------------------
Dim lose2Col As Long       '赔2数据开始列
Dim caliCol As Long       '凯利数据开始列   威廉希尔对应BFW，BETS365 对应BFB，澳门数据对应于 BFM


Dim id As String      '字典ID
Dim itemVal As Long   '字典值

Dim i, k As Integer
Dim srcCol As Long
Dim srcRow As Long

Dim multi1 As Long
Dim off1 As Long
Dim cond1



Dim tgtVal As Double
Dim srcVal As Double

Dim val1 As String   '赔变数据
Dim val2 As String   '凯同数据
Dim val3 As String
Dim val4 As String


'取出相应的字典项值

lose2Col = CLng(dict1.Item("LOSE2_COL"))
srcCol = lose2Col
srcRow = j + 2  '取即时值2

'取出大类乘积数，默认必须有
multi1 = InStr("WBMLE", class) - 1
If multi1 < 0 Then     'class填写错误，退出程序
    Exit Sub
End If
off1 = 40

'取出相应类别的凯利数据对应的数据开始列


If InStr("WBM", class) > 0 Then
    id = "BF" + class + "_COL"
    If dict1.exists(id) Then
        caliCol = CLng(dict1.Item(id))
    Else
        caliCol = -1
    End If
End If

If srcdata(j, 11) <> "" Then
    If CStr(srcdata(j, 11)) = "3" Then
        k = 0
    ElseIf CStr(srcdata(j, 11)) = "1" Then
        k = 1
    ElseIf CStr(srcdata(j, 11)) = "0" Then
        k = 2
    Else
        k = -1
    End If
End If

'处理lose2数据
For i = 0 To 2
      tgtVal = 0
    srcVal = 0
    If tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + i) <> "" Then tgtVal = CDbl(tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + i))

        
        
    If k = i Then
        If IsNull(srcdata(srcRow, srcCol + i)) Or srcdata(srcRow, srcCol + i) = "" Then srcVal = 0 Else srcVal = CDbl(srcdata(srcRow, srcCol + i))
        
        If great = True Then
            If tgtVal > srcVal Then cond1 = True Else cond1 = False
        ElseIf tgtVal < srcVal Then cond1 = True
        Else: cond1 = False
        End If
        
        If cond1 And srcVal > 0 Then
            tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + i) = srcVal
        ElseIf tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + i) = "" And Not IsNull(srcdata(srcRow, srcCol + i)) And srcVal > 0 Then
            tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + i) = srcVal
        End If
    End If
Next

'处理赔变数据
val1 = srcdata(row1 + 2, lose2Col + 3)   '基准行的赔2即时值2数据
val2 = srcdata(srcRow, lose2Col + 3)   '比较行的赔2即时值2数据
val3 = tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + 5)
If val1 <> "" And val1 = val2 Then
    tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + 4) = val1
    val4 = Trim(CStr(srcdata(srcRow, 11)))
    If val4 <> "" And InStr(val3, val4) = 0 Then '比赛结果
        tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + 5) = val3 + val4
    End If
End If

'处理凯同数据
If caliCol <> -1 Then
    val1 = srcdata(row1 + 2, caliCol + 3)   '基准行的凯利指数的即时值2数据
    val2 = srcdata(srcRow, caliCol + 3)   '比较行的凯利指数的即时值2数据
    val3 = tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + 7)
   If val1 <> "" And val1 = val2 Then
        tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + 6) = val1
        val4 = Trim(CStr(srcdata(srcRow, 11)))
        If val4 <> "" And InStr(val3, val4) = 0 Then '比赛结果
            tgtdata(tmpCol + multi1 * off1 + multi2 * off2 + 7) = val3 + val4
        End If
    End If
End If

End Sub

