Attribute VB_Name = "dataview"
Sub 查看当期数据(ByRef control As Office.IRibbonControl)
    Dim wkSheet As Worksheet
    Dim pasteSheet As Worksheet
    Dim cnt As Long
    Dim col As Integer
    Dim str As String
    Dim mode As String
    Dim okid As Integer
    Dim ctrls As Object
    
    
    Set wkSheet = ActiveWorkbook.Sheets("综合数据")
    cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
    col = wkSheet.UsedRange.Columns(wkSheet.UsedRange.Columns.Count).Column
    Set pasteSheet = ActiveWorkbook.Sheets("当期数据")
    
    str = "=" + CStr(wkSheet.Cells(1, 9))
    
    
    Call 初始化一般字典(dataColDict, wkSheet, 4, 0, 1, False)
    okid = CInt(dataColDict.Item("OKID"))
    
    
    mode = CStr(wkSheet.Cells(1, 2))
    Set ctrls = Application.CommandBars("彩票分析").Controls
        If ctrls(6).Caption = "查看全部数据" Then
            wkSheet.Range(Cells(3, 1), Cells(cnt, col)).AutoFilter Field:=9
            ctrls(6).Caption = "查看【" + CStr(wkSheet.Cells(1, 9)) + "】期"
            wkSheet.Range(Cells(3, 1), Cells(cnt, col)).AutoFilter Field:=okid
            wkSheet.Cells(cnt, 1).Select
        Else
            pasteSheet.Cells.ClearContents
            With wkSheet.Range(Cells(3, 1), Cells(cnt, col))
                .AutoFilter Field:=okid, Criteria1:=str
                '.SpecialCells(xlCellTypeVisible).Copy
                'pasteSheet.[a2]
            End With
            'pasteSheet.Range("A2").PasteSpecial
            ctrls(6).Caption = "查看全部数据"
            Call 处理当期数据
        End If
    Set wkSheet = Nothing
    Set pasteSheet = Nothing
    
End Sub

Sub 查看全部信息(ByRef control As Office.IRibbonControl)
Dim wkSheet As Worksheet
Dim cnt, col
Dim ctrls As Object

Set wkSheet = ActiveWorkbook.Sheets("综合数据")
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
col = wkSheet.UsedRange.Columns(wkSheet.UsedRange.Columns.Count).Column

Set ctrls = Application.CommandBars("彩票分析").Controls
If ctrls(7).Caption = "查看赛事信息" Then
    Call 显示筛选全集数据
    ctrls(7).Caption = "查看全部赛事"

Else
    wkSheet.Range(Cells(3, 1), Cells(cnt, col)).AutoFilter Field:=9
    ctrls(7).Caption = "查看赛事信息"
    wkSheet.Cells(cnt, 1).Select
End If
End Sub

