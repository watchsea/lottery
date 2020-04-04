Attribute VB_Name = "comm"
Public Dict As Object   '���ݲ����ֵ�
Public leagueDict As Object   '�����ֵ�
Public dataColDict As Object '���ݴ��λ���ֵ�
Public teamUniDict As Object   '���ͳһ�ֵ�


Sub ��ʼ��һ���ֵ�(tempDict As Object, paraSheet As Worksheet, idCol As Long, valCol As Long, Optional bgRow As Long = 1, Optional ColOrRow As Boolean = True)
'��ָ��SHEETҳ�еĵ�ָ�����������γ��ֵ�
'idCol:  ColOrRowΪtrueʱ�����������У�Ϊfalseʱ���������ڵ���
'valCol�� ColOrRowΪtrueʱ������ֵ�����У����valColΪ0�������кţ���Ϊfalseʱ������ֵ�����У����valColΪ0���������к�
'bgRow��������ʼ��
'ColOrRow:���л���������֯�ֵ�,Ĭ��Ϊtrue,���н�����֯; false����������֯

Dim itemId, itemVal
Dim dcnt
Dim cnt

On Error Resume Next  '�����������ִ����һ��
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
'Set paraSheet = ThisWorkbook.Sheets(sheetName)


'��ʼ������
If ColOrRow Then
    cnt = paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row
    For i = bgRow To cnt
        If paraSheet.Cells(i, idCol) <> "" Then
            itemId = paraSheet.Cells(i, idCol).Value
            If valCol > 0 Then    '�ֵ����ֵ�кŴ��ڣ��Ӷ�Ӧ������ȡֵ
                itemVal = CStr(paraSheet.Cells(i, valCol).Value)
            Else                 '�ֵ����ֵ�к�С�ڵ���0���������Ӧ�к�
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
    '���н�����֯��ĳһ�е�����Ϊ��������һ�е�����Ϊ��ֵ
    cnt = paraSheet.UsedRange.Columns(paraSheet.UsedRange.Columns.Count).Column
    For i = bgRow To cnt
        If paraSheet.Cells(idCol, i) <> "" Then
            itemId = paraSheet.Cells(idCol, i).Value
            If valCol > 0 Then    '�ֵ����ֵ�кŴ��ڣ��Ӷ�Ӧ������ȡֵ
                itemVal = CStr(paraSheet.Cells(valCol, i).Value)
            Else                 '�ֵ����ֵ�к�С�ڵ���0���������Ӧ�к�
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




Sub ��ʼ���ֵ�(tempDict As Object, sheetName As String, Optional bgRow As Long = 2, Optional keyCol As Integer = 1, Optional valCol As Integer = 3)
'--------------------------------------------------------
'������
'     tempDict�����������ֵ�
'     sheetName�����ݱ����ҳ������
'     bgRow��    ����ҳ��������ʼ��
'     keyCol��   �������ڵ��к�
'     valCol��   ��ֵ���ڵ��к�
'     add by ljqu 2016.5.8,  ���Ӳ��� bgRow,keyCol,valCol
'-------------------------------------------------------
Dim itemId, itemVal
Dim dcnt
Dim cnt

On Error Resume Next  '�����������ִ����һ��
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
Set paraSheet = ActiveWorkbook.Sheets(sheetName)


'��ʼ������

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


Function �����ۺ������ֵ�(tempDict As Object, dataArr, col1 As Integer, col2 As Integer, Optional dataType As String = "��ʼֵ")
'------------------------------------------------------------
' tempDict �������ݵ��ֵ�
' dataArr  ���ڻ�ȡ�ֵ����ݵ�����
' col1     �ֵ��ʶ��Ӧ��������
' col2     �ֵ�ֵ��Ӧ��������
' dataType :
'------------------------------------------------------------
Dim itemId, itemVal
Dim dcnt
Dim cnt




On Error Resume Next  '�����������ִ����һ��
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
    �����ۺ������ֵ� = True
Else
    �����ۺ������ֵ� = False
End If
End Function




Function BytesToBstr(strBody, CodeBase)         'ʹ��Adodb.Stream������ȡ�ַ���
    Dim objStream
    On Error Resume Next
    Set objStream = CreateObject("Adodb.Stream")
    With objStream
        .Type = 1                               '������
        .mode = 3                               '��д
        .Open
        .Write strBody                          '����������д��Adodb.Stream�����ڲ�
        .Position = 0                           'λ����ʼΪ0
        .Type = 2                               '�ַ���
        .Charset = CodeBase                     '���ݵı����ʽ
        BytesToBstr = .ReadText                 '�õ��ַ���
    End With
    objStream.Close
    Set objStream = Nothing
    If Err.Number <> 0 Then BytesToBstr = ""
    On Error GoTo 0
End Function


'��UTF-8ת��Ϊ���֣�����JS
Function UTF8toChineseCharacters(szInput)
    Dim js As Object
    Set js = CreateObjectx86("MSScriptControl.ScriptControl")
    js.Language = "JavaScript"
    js.AddCode "function decode(str){return unescape(str.replace(/\\u/g,'%u'));}"
    UTF8toChineseCharacters = js.eval("decode('" & szInput & "')")
End Function

'��JSON��ȡ��Ŀֵ
Sub getItemfromJson(aa, bb As Object)
Dim x
Dim s
     Set x = CreateObjectx86("MSScriptControl.ScriptControl")
         x.Language = "JScript"
     s = "function j(s) { return eval('(' + s + ')'); }"
       x.AddCode s
     Set bb = x.Run("j", aa)
     Set x = Nothing
End Sub

Sub �����鹹���ֵ�(tempDict As Object, dataArr, idCol As Long, valCol As Long, Optional bgRow As Long = 1)
'��ָ��SHEETҳ�еĵ�ָ�����������γ��ֵ�
'idCol:  ���������У�
'valCol�� ����ֵ�����У����valColΪ0�������кţ�
'bgRow��������ʼ��

Dim itemId, itemVal
Dim dcnt
Dim cnt

On Error Resume Next  '�����������ִ����һ��
dcnt = tempDict.Count
If IsEmpty(dcnt) Then
    Set tempDict = CreateObject("Scripting.Dictionary")
End If

'fetch the parameters
'Set paraSheet = ThisWorkbook.Sheets(sheetName)


'��ʼ������

    cnt = UBound(dataArr)   'paraSheet.UsedRange.Rows(paraSheet.UsedRange.Rows.Count).row
    For i = bgRow To cnt
        If dataArr(i, idCol) <> "" Then
            itemId = dataArr(i, idCol)
            If valCol > 0 Then    '�ֵ����ֵ�кŴ��ڣ��Ӷ�Ӧ������ȡֵ
                itemVal = CStr(dataArr(i, valCol))
            Else                 '�ֵ����ֵ�к�С�ڵ���0���������Ӧ�к�
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


'����ǰ����Microsoft Visual Basic for Application Extensibility 5.3������ѡ�����ζ�VBA���̷���
Sub ExportAllVBC()
    Dim ExportPath As String, ExtendName As String
    Dim vbc As VBComponent
    ExportPath = ThisWorkbook.path
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbc.Type
        Case vbext_ct_ClassModule, vbext_ct_Document '�������Ϊ��ģ�顢EXCEL����
            ExtendName = ".cls" '���õ����ļ�����չ��
        Case vbext_ct_MSForm '�������Ϊ����
            ExtendName = ".frm"
        Case vbext_ct_StdModule '�������Ϊģ��ʱ
            ExtendName = ".bas"
        End Select
        If ExtendName <> "" Then vbc.Export ExportPath & "\code\" & vbc.Name & ExtendName
    Next
End Sub

'�������еĽű�
Sub ImportAllVBC()
    Dim theMod As VBIDE.VBComponent
    For Each theMod In ThisWorkbook.VBProject.VBComponents
        With theMod.CodeModule
            .AddFromFile "the" & .Parent.Name & ".bas"
        End With
    Next
End Sub




Function CreateObjectx86(Optional sProgID, Optional bClose = False)
    Static oWnd As Object
    Dim bRunning As Boolean
    #If Win64 Then
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        If bClose Then
            If bRunning Then oWnd.Close
            Exit Function
        End If
        If Not bRunning Then
            Set oWnd = CreateWindow()
            oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID): End Function", "VBScript"
        End If
        Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
    #Else
        Set CreateObjectx86 = CreateObject("MSScriptControl.ScriptControl")
    #End If
End Function



Function CreateWindow()
    Dim sSignature, oShellWnd, oProc
    On Error Resume Next
    sSignature = Left(CreateObject("Scriptlet.TypeLib").GUID, 38)
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""about:<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
            Set CreateWindow = oShellWnd.GetProperty(sSignature)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        Next
    Loop
End Function



Function getUnixTime()  '��ȡUnixʱ���
    getUnixTime = DateDiff("s", "01/01/1970 00:00:00", Now())
End Function

Function getdateTime(unixTime As Long) 'UNIXʱ���ת����ʱ��
    getdateTime = DateAdd("s", unixTime, "01/01/1970 00:00:00")
End Function
