Attribute VB_Name = "comm"
Public Dict As Object   '数据参数字典
Public leagueDict As Object   '联赛字典
Public dataColDict As Object '数据存放位置字典


Sub 初始化一般字典(tempDict As Object, paraSheet As Worksheet, idCol As Long, valCol As Long, Optional bgRow As Long = 1, Optional ColOrRow As Boolean = True)
'对指定SHEET页中的的指定两列数据形成字典
'idCol:  ColOrRow为true时，主键所在列；为false时，主键所在的行
'valCol： ColOrRow为true时，主键值所在列，如果valCol为0，则填行号；若为false时，主键值所在行，如果valCol为0，则填入列号
'bgRow：数据起始行
'ColOrRow:按行还是列来组织字典,默认为true,按列进行组织; false：按行来组织

Dim itemId, itemVal
Dim dcnt
Dim cnt

On Error Resume Next  '遇到错误继续执行下一行
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
'Set paraSheet = ThisWorkbook.Sheets(sheetName)


'初始化参数
If ColOrRow Then
    cnt = paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row
    For i = bgRow To cnt
        If paraSheet.Cells(i, idCol) <> "" Then
            itemId = paraSheet.Cells(i, idCol).Value
            If valCol > 0 Then    '字典项的值列号大于，从对应于列中取值
                itemVal = CStr(paraSheet.Cells(i, valCol).Value)
            Else                 '字典项的值列号小于等于0，则填入对应行号
                itemVal = i
            End If
            
            If tempDict.exists(itemId) Then
                tempDict.Item(itemId) = itemVal
            Else
                tempDict.Add itemId, itemVal
            End If
        End If
    Next
Else
    '按行进行组织，某一行的数据为主键，另一行的数据为键值
    cnt = paraSheet.UsedRange.Columns(paraSheet.UsedRange.Columns.Count).Column
    For i = bgRow To cnt
        If paraSheet.Cells(idCol, i) <> "" Then
            itemId = paraSheet.Cells(idCol, i).Value
            If valCol > 0 Then    '字典项的值列号大于，从对应于列中取值
                itemVal = CStr(paraSheet.Cells(valCol, i).Value)
            Else                 '字典项的值列号小于等于0，则填入对应行号
                itemVal = i
            End If
            
            If tempDict.exists(itemId) Then
                tempDict.Item(itemId) = itemVal
            Else
                tempDict.Add itemId, itemVal
            End If
        End If
    Next

End If
End Sub




Sub 初始化字典(tempDict As Object, sheetName As String, Optional bgRow As Long = 2, Optional keyCol As Integer = 1, Optional valCol As Integer = 3)
'--------------------------------------------------------
'参数：
'     tempDict：待建立的字典
'     sheetName：数据保存的页面名称
'     bgRow：    数据页中数据起始行
'     keyCol：   主键所在的列号
'     valCol：   键值所在的列号
'     add by ljqu 2016.5.8,  增加参数 bgRow,keyCol,valCol
'-------------------------------------------------------
Dim itemId, itemVal
Dim dcnt
Dim cnt

On Error Resume Next  '遇到错误继续执行下一行
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
Set paraSheet = ActiveWorkbook.Sheets(sheetName)


'初始化参数

cnt = paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row
For i = bgRow To cnt
    If paraSheet.Cells(i, keyCol) <> "" Then
        itemId = paraSheet.Cells(i, keyCol).Value
        itemVal = paraSheet.Cells(i, valCol).Value
        
        If tempDict.exists(itemId) Then
            tempDict.Item(itemId) = itemVal
        Else
            tempDict.Add itemId, itemVal
        End If
    End If
Next
End Sub


Function 载入综合数据字典(tempDict As Object, dataArr, col1 As Integer, col2 As Integer, Optional dataType As String = "初始值")
'------------------------------------------------------------
' tempDict 保存数据的字典
' dataArr  用于获取字典数据的数组
' col1     字典标识对应的数据列
' col2     字典值对应的数据列
' dataType :
'------------------------------------------------------------
Dim itemId, itemVal
Dim dcnt
Dim cnt




On Error Resume Next  '遇到错误继续执行下一行
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

If IsArray(dataArr) Then
    cnt = UBound(dataArr)
    For i = 1 To cnt
        If dataArr(i, 6) = dataType And dataArr(i, col1) <> "" And dataArr(i, col2) <> "" Then
            itemId = dataArr(i, col1)
            itemVal = dataArr(i, col2)
            
            If tempDict.exists(itemId) Then
                tempDict.Item(itemId) = itemVal
            Else
                tempDict.Add itemId, itemVal
            End If
        End If
    Next
    载入综合数据字典 = True
Else
    载入综合数据字典 = False
End If
End Function




Function BytesToBstr(strBody, CodeBase)         '使用Adodb.Stream对象提取字符串
    Dim objStream
    On Error Resume Next
    Set objStream = CreateObject("Adodb.Stream")
    With objStream
        .Type = 1                               '二进制
        .mode = 3                               '读写
        .Open
        .Write strBody                          '二进制数组写入Adodb.Stream对象内部
        .Position = 0                           '位置起始为0
        .Type = 2                               '字符串
        .Charset = CodeBase                     '数据的编码格式
        BytesToBstr = .ReadText                 '得到字符串
    End With
    objStream.Close
    Set objStream = Nothing
    If Err.Number <> 0 Then BytesToBstr = ""
    On Error GoTo 0
End Function


'将UTF-8转换为汉字：调用JS
Function UTF8toChineseCharacters(szInput)
    Dim js As Object
    Set js = CreateObject("MSScriptControl.ScriptControl")
    js.Language = "JavaScript"
    js.AddCode "function decode(str){return unescape(str.replace(/\\u/g,'%u'));}"
    UTF8toChineseCharacters = js.Eval("decode('" & szInput & "')")
End Function

'从JSON中取项目值
Sub getItemfromJson(aa, bb As Object)
Dim x
Dim s
     Set x = CreateObject("ScriptControl")
         x.Language = "JScript"
     s = "function j(s) { return eval('(' + s + ')'); }"
       x.AddCode s
     Set bb = x.Run("j", aa)
     Set x = Nothing
End Sub

Sub 从数组构建字典(tempDict As Object, dataArr, idCol As Long, valCol As Long, Optional bgRow As Long = 1)
'对指定SHEET页中的的指定两列数据形成字典
'idCol:  主键所在列；
'valCol： 主键值所在列，如果valCol为0，则填行号；
'bgRow：数据起始行

Dim itemId, itemVal
Dim dcnt
Dim cnt

On Error Resume Next  '遇到错误继续执行下一行
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
'Set paraSheet = ThisWorkbook.Sheets(sheetName)


'初始化参数

    cnt = UBound(dataArr)   'paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row
    For i = bgRow To cnt
        If dataArr(i, idCol) <> "" Then
            itemId = dataArr(i, idCol)
            If valCol > 0 Then    '字典项的值列号大于，从对应于列中取值
                itemVal = CStr(dataArr(i, valCol))
            Else                 '字典项的值列号小于等于0，则填入对应行号
                itemVal = i
            End If
            
            If tempDict.exists(itemId) Then
                tempDict.Item(itemId) = itemVal
            Else
                tempDict.Add itemId, itemVal
            End If
        End If
    Next

End Sub

