Attribute VB_Name = "common"
Public Dict As Object   '数据参数字典
Public LeagueDict As Object   '联赛字典
Public DataColDict As Object '数据存放位置字典

Sub 初始化一般字典(tempDict As Object, paraSheet As Worksheet, idCol As Long, valCol As Long, Optional bgrow As Long = 1, Optional ColOrRow As Boolean = True)
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
    For i = bgrow To cnt
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
    For i = bgrow To cnt
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




Sub 初始化字典(tempDict As Object, sheetName As String)
Dim itemId, itemVal
Dim dcnt
Dim cnt

On Error Resume Next  '遇到错误继续执行下一行
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
Set paraSheet = ThisWorkbook.Sheets(sheetName)


'初始化参数

cnt = paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row
For i = 2 To cnt
    If paraSheet.Cells(i, 1) <> "" Then
        itemId = paraSheet.Cells(i, 1).Value
        itemVal = paraSheet.Cells(i, 3).Value
        
        If tempDict.exists(itemId) Then
            tempDict.Item(itemId) = itemVal
        Else
            tempDict.Add itemId, itemVal
        End If
    End If
Next
End Sub



Sub 初始化球队字典(tempDict As Object, sheetName As String)
Dim itemId, itemVal
Dim dcnt
Dim cnt

On Error Resume Next  '遇到错误继续执行下一行
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
Set paraSheet = ThisWorkbook.Sheets(sheetName)


'初始化参数

cnt = paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row
For i = 2 To cnt
    If paraSheet.Cells(i, 1) <> "" And paraSheet.Cells(i, 3) <> "" And paraSheet.Cells(i, 5) <> "" Then
        itemId = paraSheet.Cells(i, 1) & paraSheet.Cells(i, 3) & paraSheet.Cells(i, 5)
        itemVal = i
        
        If tempDict.exists(itemId) Then
            tempDict.Item(itemId) = itemVal
        Else
            tempDict.Add itemId, itemVal
        End If
    End If
Next
End Sub

Sub 初始化球队实力字典(tempDict As Object, sheetName As String, len1)
'len1表示取的数据长度，if len1=0 从头开始取，否则从总数-len1开始取len1个数据
Dim itemId, itemVal
Dim dcnt
Dim cnt, bgrow

On Error Resume Next  '遇到错误继续执行下一行
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
Set paraSheet = ThisWorkbook.Sheets(sheetName)


'初始化参数
'主键：联赛ID+球队ID+赛季+轮次
cnt = paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row

If len1 = 0 Then
    bgrow = 2
Else
    bgrow = cnt - len1 + 1
End If


For i = bgrow To cnt
    If paraSheet.Cells(i, 1) <> "" And paraSheet.Cells(i, 3) <> "" And paraSheet.Cells(i, 5) <> "" And paraSheet.Cells(i, 6) <> "" Then
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
    js.addcode "function decode(str){return unescape(str.replace(/\\u/g,'%u'));}"
    UTF8toChineseCharacters = js.Eval("decode('" & szInput & "')")
End Function

'从JSON中取项目值
Sub getItemfromJson(aa, bb As Object)
Dim x
Dim s
     Set x = CreateObject("ScriptControl")
         x.Language = "JScript"
     s = "function j(s) { return eval('(' + s + ')'); }"
     's = "function j(s) { return eval(s);}"
       x.addcode s
     Set bb = x.Run("j", aa)
     Set x = Nothing
End Sub




Function openFile(fileType As Integer, Optional filePath As String = "", Optional fileName As String = "")

'******************************************************************************
'* Purpose:  利用文件打开对话框，选择要打开的文件
'* Author:   ljqu
'* create date: 12/20/2010
'* Category:
'* Version:  1.0
'* Company:
'* Param:
'*       fileType:1:TXT,2:EXCEL
'*
'******************************************************************************
    Dim boolResult, flPath, flName
    
    'Set objDialog = CreateObject("UserAccounts.CommonDialog")
    
    Dim objDialog As FileDialog
    
    Set objDialog = Application.FileDialog(msoFileDialogOpen)
    objDialog.Filters.Clear
    
    If (fileType = 1) Then
        objDialog.Filters.Add "所有文件", "*.*"
        objDialog.Filters.Add "文本文件", "*.txt"
    ElseIf (fileType = 2) Then
        objDialog.Filters.Add "Excel 97-2003(*.xls)", "*.xls"
        objDialog.Filters.Add "Excel 2007(*.xlsx)", "*.xlsx"
        objDialog.Filters.Add "所有文件", "*.*"
    ElseIf (fileType = 3) Then
        objDialog.Filters.Add "所有文件", "*.*"
        objDialog.Filters.Add "Excel 宏文件(*.xlsm)", "*.xlsm"
    End If
    objDialog.FilterIndex = 2
    If filePath = "" Then
        flPath = ThisWorkbook.path
    Else
        flPath = filePath
    End If
    
    If fileName <> "" Then
        flName = fileName
    End If
    
    objDialog.InitialFileName = flPath + "\" + flName
    
    boolResult = objDialog.Show
    
    If boolResult = 0 Then
       openFile = "Empty"
    Else
       'Output ("You choose: " & objDialog.Filename)
       openFile = objDialog.SelectedItems(1)
    End If
End Function


Sub 加载数据到内存(sheetName As String, data, bgrow As Long, Optional bgCol As Long = 1)
'*****************************************************************
'  把Sheet页中的数据加载内存数组中
'  sheetName： 待加载的数据Sheet页名称
'  data:    存储数据的数组
'  bgRow:   数据开始的行号
'  boCol:   数据开始的列号,默认为从第1列开始
'*****************************************************************
Dim wkSheet As Worksheet
Dim row1, col1

Dim i, j

Set wkSheet = ThisWorkbook.Sheets(sheetName)
row1 = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
col1 = wkSheet.UsedRange.Columns(wkSheet.UsedRange.Columns.Count).Column
ReDim data(row1, col1)
For i = bgrow To row1
    data(i, 0) = i    '保留行号
    For j = bgCol To col1
        data(i, j) = wkSheet.Cells(i, j)
    Next
Next

Set wkSheet = Nothing

End Sub

