Attribute VB_Name = "dataview"
Sub �鿴��������(ByRef control As Office.IRibbonControl)
    Dim wkSheet As Worksheet
    Dim pasteSheet As Worksheet
    Dim cnt As Long
    Dim col As Integer
    Dim str As String
    Dim mode As String
    Dim okid As Integer
    Dim ctrls As Object
    
    
    Set wkSheet = ActiveWorkbook.Sheets("�ۺ�����")
    cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
    col = wkSheet.UsedRange.Columns(wkSheet.UsedRange.Columns.Count).Column
    Set pasteSheet = ActiveWorkbook.Sheets("��������")
    
    str = "=" + CStr(wkSheet.Cells(1, 9))
    
    
    Call ��ʼ��һ���ֵ�(dataColDict, wkSheet, 4, 0, 1, False)
    okid = CInt(dataColDict.Item("OKID"))
    
    
    mode = CStr(wkSheet.Cells(1, 2))
    Set ctrls = Application.CommandBars("��Ʊ����").Controls
        If ctrls(6).Caption = "�鿴ȫ������" Then
            wkSheet.Range(Cells(3, 1), Cells(cnt, col)).AutoFilter Field:=9
            ctrls(6).Caption = "�鿴��" + CStr(wkSheet.Cells(1, 9)) + "����"
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
            ctrls(6).Caption = "�鿴ȫ������"
            Call ����������
        End If
    Set wkSheet = Nothing
    Set pasteSheet = Nothing
    
End Sub

Sub �鿴ȫ����Ϣ(ByRef control As Office.IRibbonControl)
Dim wkSheet As Worksheet
Dim cnt, col
Dim ctrls As Object

Set wkSheet = ActiveWorkbook.Sheets("�ۺ�����")
cnt = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row
col = wkSheet.UsedRange.Columns(wkSheet.UsedRange.Columns.Count).Column

Set ctrls = Application.CommandBars("��Ʊ����").Controls
If ctrls(7).Caption = "�鿴������Ϣ" Then
    Call ��ʾɸѡȫ������
    ctrls(7).Caption = "�鿴ȫ������"

Else
    wkSheet.Range(Cells(3, 1), Cells(cnt, col)).AutoFilter Field:=9
    ctrls(7).Caption = "�鿴������Ϣ"
    wkSheet.Cells(cnt, 1).Select
End If
End Sub

