Attribute VB_Name = "histdata"
Option Explicit
#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub 历史数据加载(ByRef control As Office.IRibbonControl)
    Dim begindate As Date
    Dim enddate As Date
    
    
    Call 初始化字典(Dict, "Param")
    Call 初始化字典(leagueDict, "01赛事")
    
    begindate = CDate(Dict.Item("BEGIN_DATE"))
    enddate = CDate(Dict.Item("END_DATE"))
    If begindate > enddate Then
        MsgBox ("开始日期大于结束日期,请重新输入！")
        Exit Sub
    End If
    
    If MsgBox("日期：" + CStr(begindate) + " ----- " + CStr(enddate) + Chr(13) + "如要修改起止日期，请进入参数页进行修改。", vbOKCancel, "确认信息") = vbCancel Then
        Exit Sub
    End If
    
    
    If 球探网历史数据载入("球探网(W)", "id=115&company=威廉希尔(英国)", begindate, enddate) Then
        
        Call 球探网BF数据载入
        Call 球探网赛事积分数据载入
        
        Call 澳客网必发盈亏(begindate, enddate)
        Call 澳客网胜负指数(begindate, enddate)
        Call 澳客网盘口评测(begindate, enddate)
        Call 澳客网凯利指数(begindate, enddate)
        
        
       MsgBox ("历史数据加载完毕！请往下点击【初始】按钮。")
    End If

End Sub



Function 球探网历史数据载入(sheetName As String, ids, begindate As Date, enddate As Date)
'------------------------------------------------------------------
'ids：相关公司对应的进入参数
'beginDate：历史数据的开始日期
'endDate:历史数据的结束日期
'------------------------------------------------------------------
Dim IE As Object
Dim doc As Object
Dim k As Long
Dim i As Long
Dim j As Long
Dim rowcnt As Long
Dim colCnt As Long
Dim data1()

Dim itemId
Dim col As Long

Dim tt As Object
Dim tt1 As Object
Dim tt2, tt3
Dim URL
Dim wkSheet As Worksheet
Dim isFirst As Boolean
Dim shCounter As Integer    '工作表记录指针
Dim calDate As Date



'清除页面内容
Set wkSheet = ActiveWorkbook.Worksheets(sheetName)
wkSheet.Cells.ClearContents
isFirst = True
shCounter = 1             '第一次从cells(1,1)开始

球探网历史数据载入 = True

'ids = "1014941"   '"987108"
calDate = begindate
Do While calDate <= enddate

    URL = "http://1x2.win007.com/Companyhistory.aspx?type=1&"    'type=1表示的指定日期的数据
    URL = URL + ids
    
    URL = URL + "&matchdate=" + CStr(calDate)
    
    Set IE = UserForm1.WebBrowser1
    With IE
      .Navigate URL '网址
      Do Until .ReadyState = 4
        DoEvents
      Loop
      Set doc = .document
    End With
    'Application.ScreenUpdating = False
    
    If doc.body.innerHTML = "访问频率超出限制。" Then
        MsgBox ("今日下载历史数据已经到极限！   请隔日下载。")
        球探网历史数据载入 = False
        Set wkSheet = Nothing
        Exit Function
    End If
    Set tt = doc.getElementById("table_schedule").getElementsByTagName("tr")
    rowcnt = tt.Length - 1
    colCnt = tt(0).Cells.Length - 1
    col = 0
    ReDim data1(rowcnt, colCnt + 7)  '重置数组
    
    '读取表头
    For j = 0 To tt(0).Cells.Length - 1
        If j < 11 Then
            data1(0, j + 1) = tt(0).Cells(j).innerText
        ElseIf j = 11 Then
            data1(0, 0) = tt(0).Cells(j).innerText
        Else
            data1(0, j) = tt(0).Cells(j).innerText
        End If
    Next
    For k = 3 To 9
        data1(col, k - 3 + j) = tt(0).Cells(k).innerText
    Next
    
    For i = 1 To rowcnt    '处理数据
        tt2 = tt(i).getAttribute("name")   '取名称，即联赛的ID
        If Not IsNull(tt2) Then
            tt3 = Split(tt2, ",")(0)          '获取联赛ID
            '判断联赛ID是否在要取的联赛ID中
            If tt3 <> "" Then
                itemId = tt(i).Cells(0).innerText    '联赛
                If leagueDict.exists(itemId) Then
                    col = col + 1
                    For j = 0 To tt(i).Cells.Length - 1
                        If j = colCnt Then
                            itemId = tt(i).Cells(j).ChildNodes(0).nameProp
                            data1(col, j) = Left(itemId, Len(itemId) - 4)
                        ElseIf j < 11 Then
                            data1(col, j + 1) = tt(i).Cells(j).innerText
                        ElseIf j = 11 Then
                            data1(col, 0) = Split(tt(i).Cells(j).innerText, Chr(13))(0)
                        Else
                            data1(col, j) = tt(i).Cells(j).innerText
                        End If
                    Next
                    For k = 0 To tt(i + 1).Cells.Length - 1
                        data1(col, j + k) = tt(i + 1).Cells(k).innerText
                    Next
                End If
            End If
        End If
    Next
    
    Set IE = Nothing
    
    
    '将数据过录到SHEET页
    If isFirst Then      '第一次过录，需要写入表头
        k = 0
        isFirst = False
    Else                 '第二次过录，则不写入表头
        k = 1
    End If
    For i = k To col
        For j = 0 To colCnt + 7
            If j = 2 Then wkSheet.Cells(shCounter + i, j + 1).NumberFormatLocal = "@"
            wkSheet.Cells(shCounter + i, j + 1) = data1(i, j)
        Next
    Next
    shCounter = shCounter + col
    calDate = DateAdd("d", 1, calDate)   '增加一天
    
    '延时1秒
    Sleep 500
Loop


Set wkSheet = Nothing

End Function



Sub 澳客网必发盈亏(begindate As Date, enddate As Date)

Dim doc As Object 'MSHTML.HTMLDocument
Dim objXml As Object
Dim oDoc As Object
Dim txt As String
Dim txt1

Dim recCnt As Integer    '总记录数
Dim pageCnt As Integer   '总页数
Dim k As Integer
Dim i As Integer
Dim j As Integer
Dim t1 As String
Dim t2 As String
Dim tbody As Object
Dim ucell As Object
Dim rowcnt As Integer
Dim totcnt As Integer   '总的记录数
Dim data()
Dim WebBrowser1 As Object
Dim wkSheet As Worksheet
Dim dt As Date
Dim sdt As String

Dim URL As String, postData As String


Dim waitcnt As Integer   '延时等待次数   ,   2018.11.1  为了处理在加载过程部分页面无响应而导致的程序死循环而设计
Dim errorinfo As String   '记录被装载跳过的页数  2018.11.1  为了处理在加载过程部分页面无响应而导致的程序死循环而设计

'dt = "2014-9-27"

'清空页面，为历史数据装载做准备
Set wkSheet = ActiveWorkbook.Sheets("澳客网(1)")
wkSheet.Cells.Clear

'wkSheet.Cells(1, 1) = "期数"
wkSheet.Cells(1, 2) = "标识"
wkSheet.Cells(1, 3) = "本期编号"
wkSheet.Cells(1, 4) = "联赛"
wkSheet.Cells(1, 5) = "日期"
wkSheet.Cells(1, 6) = "时间"
wkSheet.Cells(1, 7) = "主队"
wkSheet.Cells(1, 8) = "比分"
wkSheet.Cells(1, 9) = "客队"
wkSheet.Cells(1, 10) = "买家挂牌-主胜"
wkSheet.Cells(1, 11) = "平局"
wkSheet.Cells(1, 12) = "客胜"
wkSheet.Cells(1, 13) = "买家价位-主胜"
wkSheet.Cells(1, 14) = "平局"
wkSheet.Cells(1, 15) = "客胜"

wkSheet.Cells(1, 16) = "卖家挂牌-主胜"
wkSheet.Cells(1, 17) = "平局"
wkSheet.Cells(1, 18) = "客胜"
wkSheet.Cells(1, 19) = "卖家价位-主胜"
wkSheet.Cells(1, 20) = "平局"
wkSheet.Cells(1, 21) = "客胜"


wkSheet.Cells(1, 22) = "总成交额-主胜"
wkSheet.Cells(1, 23) = "平局"
wkSheet.Cells(1, 24) = "客胜"
wkSheet.Cells(1, 25) = "冷热指数-主胜"
wkSheet.Cells(1, 26) = "平局"
wkSheet.Cells(1, 27) = "客胜"


wkSheet.Cells(1, 28) = "市场指数-主胜"
wkSheet.Cells(1, 29) = "平局"
wkSheet.Cells(1, 30) = "客胜"
wkSheet.Cells(1, 31) = "必发赔率-主胜"
wkSheet.Cells(1, 32) = "平局"
wkSheet.Cells(1, 33) = "客胜"


wkSheet.Cells(1, 34) = "必发比例-主胜"
wkSheet.Cells(1, 35) = "平局"
wkSheet.Cells(1, 36) = "客胜"
wkSheet.Cells(1, 37) = "99平均-主胜"
wkSheet.Cells(1, 38) = "平局"
wkSheet.Cells(1, 39) = "客胜"


wkSheet.Cells(1, 40) = "99平均比例-主胜"
wkSheet.Cells(1, 41) = "平局"
wkSheet.Cells(1, 42) = "客胜"
wkSheet.Cells(1, 43) = "竞彩比例-主胜"
wkSheet.Cells(1, 44) = "平局"
wkSheet.Cells(1, 45) = "客胜"

wkSheet.Cells(1, 46) = "模拟盈亏-主胜"
wkSheet.Cells(1, 47) = "平局"
wkSheet.Cells(1, 48) = "客胜"


wkSheet.Columns("C:C").NumberFormatLocal = "@"
wkSheet.Columns("E:E").NumberFormatLocal = "yyyy/m/d"
wkSheet.Columns("H:H").NumberFormatLocal = "@"
totcnt = 1

Set WebBrowser1 = UserForm1.WebBrowser1
Set objXml = CreateObject("MSXML2.XMLHTTP")
Set oDoc = CreateObject("htmlfile")

dt = begindate
'Do While dt <= dt    '无法处理历史数据，所以只一次循环即可，将enddate改为begindate
    sdt = CStr(dt)
    sdt = Replace(sdt, "/", "-")
    'WebBrowser1.Navigate "http://www.okooo.com/jingcai/shuju/betfa/" + sdt   'delete 2018.10.17
    WebBrowser1.Navigate "http://www.okooo.com/danchang/shuju/betfa/"
'    WebBrowser1.Refresh
    Do Until WebBrowser1.ReadyState = 4
        DoEvents
    Loop
    Set doc = WebBrowser1.document
    Set tbody = doc.getElementsByTagName("div")
    
    
    '***********************************************
    '   由于在实际的页面访问中会出现500错误
    '   因而先将pageCnt和recCnt 置为0
    '   2016.9.4   add by ljqu
    '***********************************************
    recCnt = 0
    pageCnt = 0
    
    

    For Each ucell In tbody
        If ucell.className = "pagination" Then    '获取页数信息
            'MsgBox ("页码")
            If ucell.innerText = "暂时没有添加记录" Then   '没有内容，则执行下一天
                recCnt = 0
                pageCnt = 0
                Exit For
            End If
            
            txt = Split(ucell.innerText, "页首页上一页")(0)
             txt1 = Split(txt, "条记录共")
             pageCnt = txt1(1)
             recCnt = Split(txt1(0), "共有")(1)
             Exit For
        End If
    Next
    
    '预处理第一页数据，获取相关参数数据
    If pageCnt > 0 Then
        '标识，期数，联赛，日期，时间，主队，比分，客队，必发赔率(主胜、平局、客胜）、99家平均
        ReDim data(recCnt, 47)
        
        '处理后续网页的参数
        '球队
        postData = "LeagueID="
        Set oDoc = doc.getElementById("csfilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        '让球数据
        postData = postData & "&HandicapNumber="
        Set oDoc = doc.getElementById("rqfilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        
        '数据
        postData = postData & "&BetDate="
        Set oDoc = doc.getElementById("datafilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        
        'MakerType  &  HasEnd
        postData = postData & "&MakerType=undefined&HasEnd=1&PageID="  '2, pageNo
        
    End If
    
    '处理所有页面的数据，2019.12.2此处略作调整，由于页面默认是查看未结束的页面，而在实际中需要取已结束的球队数据
    '因而对第一页需读取两次。
    errorinfo = ""
    rowcnt = 1
    
    i = 1
    Do
    'For i = 1 To pageCnt
        Sleep 500 * Round(Rnd, 2) + 500 * Round(Rnd, 2) + 1000 * Round(Rnd, 3)
        
        
    
        'URL = "http://www.okooo.com/danchang/shuju/betfa"
        'postData = postData & i
        URL = "http://www.okooo.com/danchang/shuju/betfa?" & postData & i
        Debug.Print URL
        WebBrowser1.Navigate URL   '"javascript:JsGoTo(" + CStr(i) + ")"     '原有分号，2015.7.27去掉“;"
        waitcnt = 0
        Do Until WebBrowser1.ReadyState = 4 Or waitcnt > 30
            Sleep 1000
            waitcnt = waitcnt + 1
            If waitcnt > 30 Then    '超过10S的等待则跳过
                errorinfo = errorinfo & i & ","
            End If
            DoEvents
        Loop
        If WebBrowser1.ReadyState = 4 Then
            Set doc = WebBrowser1.document
            
            If i = 1 Then   '重新取一次页码数据
                Set tbody = doc.getElementsByTagName("div")
                For Each ucell In tbody
                    If ucell.className = "pagination" Then    '获取页数信息
                        'MsgBox ("页码")
                        If ucell.innerText = "暂时没有添加记录" Then   '没有内容，则执行下一天
                            recCnt = 0
                            pageCnt = 0
                            Exit For
                        End If
                        
                        txt = Split(ucell.innerText, "页首页上一页")(0)
                         txt1 = Split(txt, "条记录共")
                         pageCnt = txt1(1)
                         recCnt = Split(txt1(0), "共有")(1)
                         Exit For
                    End If
                Next
                ReDim data(recCnt, 47)
            End If
            
            '处理相关资料
            rowcnt = 规范澳客网必发盈亏数据(doc, rowcnt, data, dt)
            
            
        End If
        
        
    'Next i
    i = i + 1
    Loop While i <= pageCnt
    
    
    If errorinfo <> "" Then
        MsgBox ("必发赢亏数据加载失败的页数包括(" & errorinfo & ")")
    End If
    
    '数据过录
    For i = 1 To recCnt
        For j = 0 To 47
            wkSheet.Cells(i + totcnt, j + 1) = data(i, j)
        Next
    Next
    totcnt = totcnt + recCnt
    
    dt = DateAdd("d", 1, dt)   '增加一天
    Sleep 500
'Loop

wkSheet.Cells(1, 1) = totcnt

WebBrowser1.Navigate "about:blank"
Set wkSheet = Nothing

End Sub



Function 规范澳客网必发盈亏数据(doc As Object, cnt As Integer, BFarr, pDt As Date)
'------------------------------------------------------------------------------------
'功能说明：  对澳客网必发盈亏数据进行规范化
'参数说明：
'         doc： 对应的网页对象
'         cnt：数组开始记录的下标
'         BFarr：记录数据的数组
'         pDt： 澳客网竞彩日期
'-------------------------------------------------------------------------------------
Dim divObjects As Object
Dim divObj As Object
Dim i, j, k, tt
Dim node As Object
Dim BasicInfo As Object
Dim Klinfo As Object
Dim Loc As Integer

    Loc = cnt
    Set divObjects = doc.body.All.tags("DIV")
    For i = 0 To divObjects.Length - 1
        Set divObj = divObjects(i)
        If divObj.className = "clearfix container_wrapper betfa" Then     '输出的数据
              Set node = divObj.ChildNodes
              Set BasicInfo = node(0).ChildNodes
              Set Klinfo = node(1).Rows
              
              If Loc <= UBound(BFarr) Then    '2018.10.18 add ,由于在第一次取记录后，数据会随时增长，导致读到最后一页时的数据已经发生变更，
                              '因而将所取数据限定在第一次刷新网页时的记录数。
                BFarr(Loc, 0) = pDt   'lotteryNo                 '期数
                BFarr(Loc, 1) = BasicInfo(0).innerText                   '标识
                BFarr(Loc, 2) = Right(BasicInfo(0).ChildNodes(0).innerText, 3)    '期数
                BFarr(Loc, 3) = BasicInfo(0).ChildNodes(1).innerText     '联赛
                tt = Split(BasicInfo(0).ChildNodes(2).innerText, " ")
                If DatePart("m", pDt) = 12 And Split(tt(0), "-")(0) = "01" Then
                    BFarr(Loc, 4) = CStr(DatePart("yyyy", pDt) + 1) & "-" & tt(0)  '日期
                Else
                    BFarr(Loc, 4) = CStr(DatePart("yyyy", pDt)) & "-" & tt(0)    '日期
                End If
                BFarr(Loc, 5) = tt(1)     '时间
                For j = 0 To BasicInfo(1).ChildNodes.Length - 1
                    If "SPAN" = BasicInfo(1).ChildNodes(j).nodename Then
                          BFarr(Loc, 6) = BasicInfo(1).ChildNodes(j).innerText     '主队
                    ElseIf "STRONG" = BasicInfo(1).ChildNodes(j).nodename Then
                          BFarr(Loc, 7) = BasicInfo(1).ChildNodes(j).innerText     '比分
                    ElseIf "B" = BasicInfo(1).ChildNodes(j).nodename Then
                          BFarr(Loc, 8) = BasicInfo(1).ChildNodes(j).innerText     '客队
                    End If
                Next j
                For j = 0 To Klinfo.Length - 1
                     If InStr(Klinfo(j).innerText, "主胜") > 0 Then
                          For k = 0 To 12
                          BFarr(Loc, 9 + k * 3) = Klinfo(j).Cells(k + 1).innerText
                          'BFarr(loc, 9) = Klinfo(j).Cells(8).innerText                       '必发
                          'BFarr(loc, 12) = Klinfo(j).Cells(10).innerText
                          Next k '99平均
                     End If
                     If InStr(Klinfo(j).innerText, "平局") > 0 Then
                       For k = 0 To 12
                          BFarr(Loc, 9 + k * 3 + 1) = Klinfo(j).Cells(k + 1).innerText
                          'BFarr(loc, 9) = Klinfo(j).Cells(8).innerText                       '必发
                          'BFarr(loc, 12) = Klinfo(j).Cells(10).innerText
                          Next k '99平均
                     End If
                     If InStr(Klinfo(j).innerText, "客胜") > 0 Then
                          For k = 0 To 12
                          BFarr(Loc, 9 + k * 3 + 2) = Klinfo(j).Cells(k + 1).innerText
                          'BFarr(loc, 9) = Klinfo(j).Cells(8).innerText                       '必发
                          'BFarr(loc, 12) = Klinfo(j).Cells(10).innerText
                          Next k '99平均
                     End If
                Next
                Loc = Loc + 1
             End If
        End If
    Next


Set divObjects = Nothing
Set divObj = Nothing
Set node = Nothing
Set BasicInfo = Nothing
Set Klinfo = Nothing

规范澳客网必发盈亏数据 = Loc

End Function


Function 澳客网胜负指数(begindate As Date, enddate As Date)
'------------------------------------------------------------------------------------
'功能说明：  采集澳客网的胜负指数数据数据
'
'-------------------------------------------------------------------------------------
Dim WebBrowser1 As Object
Dim wkSheet As Worksheet
Dim dt As Date
Dim sdt As String

Dim doc As Object
Dim divObjects As Object
Dim divObj As Object
Dim i, j, tt
Dim node As Object
Dim BasicInfo As Object
Dim Klinfo As Object
Dim Loc As Integer

dt = "2014-9-15"





Set wkSheet = ActiveWorkbook.Sheets("澳客网(2)")
wkSheet.Cells.Clear

'wkSheet.Cells(1, 1) = "期数"
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


Loc = 2
Set WebBrowser1 = UserForm1.WebBrowser1
dt = begindate
'Do While dt <= enddate
    sdt = CStr(dt)
    sdt = Replace(sdt, "/", "-")
    WebBrowser1.Navigate "http://www.okooo.com/danchang/shuju/zhishu/"     '+ sdt
'    WebBrowser1.Refresh
    Do Until WebBrowser1.ReadyState = 4
        DoEvents
    Loop
    Set doc = WebBrowser1.document
    
    Set divObjects = doc.body.All.tags("TABLE")
    For i = 0 To divObjects.Length - 1
        Set divObj = divObjects(i)
        If divObj.className = "magazine_table" Then     '输出的数据
              Set Klinfo = divObj.Rows
              For j = 1 To Klinfo.Length - 1            '第一行为标题，忽略
                wkSheet.Cells(Loc, 1) = dt                 '期数
                wkSheet.Cells(Loc, 2) = Klinfo(j).Cells(0).innerText + Klinfo(j).Cells(1).innerText + Klinfo(j).Cells(2).innerText 'Klinfo(j).innerText                   '标识
                wkSheet.Cells(Loc, 3) = Right(Klinfo(j).Cells(0).innerText, 3)    '编号
                wkSheet.Cells(Loc, 4) = Klinfo(j).Cells(1).innerText     '联赛
                tt = Split(Klinfo(j).Cells(2).innerText, " ")
                If DatePart("m", pDt) = 12 And Split(tt(0), "-")(0) = "01" Then
                    wkSheet.Cells(Loc, 5) = CStr(DatePart("yyyy", dt) + 1) & "-" & tt(0)  '日期
                Else
                    wkSheet.Cells(Loc, 5) = CStr(DatePart("yyyy", dt)) & "-" & tt(0)    '日期
                End If
                
                
                wkSheet.Cells(Loc, 6) = tt(1)     '时间
                
                
                wkSheet.Cells(Loc, 7) = Klinfo(j).Cells(3).innerText     '主队
                wkSheet.Cells(Loc, 8) = Klinfo(j).Cells(4).innerText     '比分
                wkSheet.Cells(Loc, 9) = Klinfo(j).Cells(5).innerText     '客队
                
                '初始值
                wkSheet.Cells(Loc, 10) = Klinfo(j).Cells(6).innerText                       '主胜
                wkSheet.Cells(Loc, 11) = Klinfo(j).Cells(7).innerText                     '平局
                wkSheet.Cells(Loc, 12) = Klinfo(j).Cells(8).innerText                     '客胜
                
                '即时值
                wkSheet.Cells(Loc, 13) = Klinfo(j).Cells(9).innerText                       '主胜
                wkSheet.Cells(Loc, 14) = Klinfo(j).Cells(10).innerText                     '平局
                wkSheet.Cells(Loc, 15) = Klinfo(j).Cells(11).innerText                     '客胜
                Loc = Loc + 1
              Next
        End If
    Next
    '清除本次的变量
    Set Klinfo = Nothing
    Set BasicInfo = Nothing
    Set divObj = Nothing
    Set divObjects = Nothing
    
    dt = DateAdd("d", 1, dt)   '增加一天
    Sleep 500
'Loop
   wkSheet.Cells(1, 1) = Loc - 1
   WebBrowser1.Navigate "about:blank"
   Set wkSheet = Nothing
   
End Function



Sub 澳客网凯利指数(begindate As Date, enddate As Date)

Dim doc As Object 'MSHTML.HTMLDocument
Dim txt As String
Dim txt1

Dim recCnt As Integer    '总记录数
Dim pageCnt As Integer   '总页数
Dim k As Integer
Dim i As Integer
Dim j As Integer
Dim t1 As String
Dim t2 As String
Dim tbody As Object
Dim ucell As Object
Dim rowcnt As Integer
Dim totcnt As Integer
Dim data()
Dim WebBrowser1 As Object
Dim wkSheet As Worksheet
Dim dt As Date
Dim sdt As String


'add by ljqu  2018.12.14
Dim URL As String, postData As String
Dim oDoc As Object
'add 2018.12.14 end

Dim waitcnt As Integer   '延时等待次数   ,   2018.11.1  为了处理在加载过程部分页面无响应而导致的程序死循环而设计
Dim errorinfo As String   '记录被装载跳过的页数  2018.11.1  为了处理在加载过程部分页面无响应而导致的程序死循环而设计


Set wkSheet = ActiveWorkbook.Sheets("澳客网(3)")
wkSheet.Cells.Clear

'wkSheet.Cells(1, 1) = "期数"
wkSheet.Cells(1, 2) = "标识"
wkSheet.Cells(1, 3) = "本期编号"
wkSheet.Cells(1, 4) = "联赛"
wkSheet.Cells(1, 5) = "日期"
wkSheet.Cells(1, 6) = "时间"
wkSheet.Cells(1, 7) = "主队"
wkSheet.Cells(1, 8) = "比分"
wkSheet.Cells(1, 9) = "客队"
wkSheet.Cells(1, 10) = "威廉希尔-主胜"
wkSheet.Cells(1, 11) = "平局"
wkSheet.Cells(1, 12) = "客胜"
wkSheet.Cells(1, 13) = "赔付率"
wkSheet.Cells(1, 14) = "Bet365-主胜"
wkSheet.Cells(1, 15) = "平局"
wkSheet.Cells(1, 16) = "客胜"
wkSheet.Cells(1, 17) = "赔付率"
wkSheet.Cells(1, 18) = "澳门-主胜"
wkSheet.Cells(1, 19) = "平局"
wkSheet.Cells(1, 20) = "客胜"
wkSheet.Cells(1, 21) = "赔付率"
wkSheet.Cells(1, 22) = "凯利指数-主胜"
wkSheet.Cells(1, 23) = "平局"
wkSheet.Cells(1, 24) = "客胜"

wkSheet.Columns("C:C").NumberFormatLocal = "@"
wkSheet.Columns("E:E").NumberFormatLocal = "yyyy/m/d"
wkSheet.Columns("H:H").NumberFormatLocal = "@"
totcnt = 1

Set WebBrowser1 = UserForm1.WebBrowser1
dt = begindate
'Do While dt <= enddate
    sdt = CStr(dt)
    sdt = Replace(sdt, "/", "-")
    WebBrowser1.Navigate "http://www.okooo.com/danchang/shuju/peilv/"  '+ sdt
'    WebBrowser1.Refresh
    Do Until WebBrowser1.ReadyState = 4
        DoEvents
    Loop
    Set doc = WebBrowser1.document
    Set tbody = doc.getElementsByTagName("div")


    '***********************************************
    '   由于在实际的页面访问中会出现500错误
    '   因而先将pageCnt和recCnt 置为0
    '   2016.9.4   add by ljqu
    '***********************************************
    recCnt = 0
    pageCnt = 0
    
    
    For Each ucell In tbody
        If ucell.className = "pagination" Then    '获取页数信息
            'MsgBox ("页码")
            If ucell.innerText = "暂时没有添加记录" Then   '没有内容，则执行下一天
                recCnt = 0
                pageCnt = 0
                Exit For
            End If
            txt = Split(ucell.innerText, "页首页上一页")(0)
             txt1 = Split(txt, "条记录共")
             pageCnt = txt1(1)
             recCnt = Split(txt1(0), "共有")(1)
             Exit For
        End If
    Next
    
    '处理第一页数据
    If pageCnt > 0 Then
        '标识，期数，联赛，日期，时间，主队，比分，客队，必发赔率(主胜、平局、客胜）、99家平均
        ReDim data(recCnt, 23)
        
        
        '处理后续页面的参数
        postData = "LeagueID="
        Set oDoc = doc.getElementById("csfilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        '让球数据
        postData = postData & "&HandicapNumber="
        Set oDoc = doc.getElementById("rqfilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        
        '数据
        postData = postData & "&BetDate="
        Set oDoc = doc.getElementById("datafilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        
        'MakerType
        Set oDoc = doc.getElementById("makerTypeObj")
        postData = postData & "&MakerType=" & oDoc.select_company
        '&  HasEnd
        postData = postData & "&HasEnd=1&PageID="  '2, pageNo
    End If

    '处理剩余页面的数据
    errorinfo = ""
    rowcnt = 1
    i = 1
    Do
    'For i = 1 To pageCnt
        Sleep 500 * Round(Rnd, 2) + 500 * Round(Rnd, 2) + 1000 * Round(Rnd, 3)
        
        URL = "http://www.okooo.com/danchang/shuju/peilv?" & postData & i
        Debug.Print URL
        WebBrowser1.Navigate URL   '"javascript:JsGoTo(" + CStr(i) + ")"     '原有分号，2015.7.27去掉“;"
        
        waitcnt = 0
        Do Until WebBrowser1.ReadyState = 4 Or waitcnt > 30
            Sleep 1000
            waitcnt = waitcnt + 1
            If waitcnt > 30 Then    '超过10S的等待则跳过
                errorinfo = errorinfo & i & ","
            End If
            DoEvents
        Loop
        If WebBrowser1.ReadyState = 4 Then
            Set doc = WebBrowser1.document
            If i = 1 Then   '重新取一次页码数据
                Set tbody = doc.getElementsByTagName("div")
                For Each ucell In tbody
                    If ucell.className = "pagination" Then    '获取页数信息
                        'MsgBox ("页码")
                        If ucell.innerText = "暂时没有添加记录" Then   '没有内容，则执行下一天
                            recCnt = 0
                            pageCnt = 0
                            Exit For
                        End If
                        
                        txt = Split(ucell.innerText, "页首页上一页")(0)
                         txt1 = Split(txt, "条记录共")
                         pageCnt = txt1(1)
                         recCnt = Split(txt1(0), "共有")(1)
                         Exit For
                    End If
                Next
                ReDim data(recCnt, 23)
            End If
            '处理相关资料
            rowcnt = 规范澳客网凯利指数(doc, rowcnt, data, dt)
        End If
    'Next i
    i = i + 1
    Loop While i <= pageCnt
    
    
    If errorinfo <> "" Then
        MsgBox ("凯利指数数据加载失败的页数包括(" & errorinfo & ")")
    End If
    
    '数据过录到EXCEL
    For i = 1 To recCnt
        For j = 0 To 23
            wkSheet.Cells(i + totcnt, j + 1) = data(i, j)
        Next
    Next
    
    totcnt = totcnt + recCnt
    
    dt = DateAdd("d", 1, dt)   '增加一天
    Sleep 150
'Loop
wkSheet.Cells(1, 1) = totcnt
    WebBrowser1.Navigate "about:blank"
    Set wkSheet = Nothing
    
End Sub


Function 规范澳客网凯利指数(doc As Object, cnt As Integer, KLarr, pDt As Date)
'------------------------------------------------------------------------------------
'功能说明：  对澳客网凯利指数数据进行规范化
'参数说明：
'         doc： 对应的网页对象
'         cnt：数组开始记录的下标
'         BFarr：记录数据的数组
'         pDt： 澳客网竞彩日期
'-------------------------------------------------------------------------------------
Dim divObjects As Object
Dim divObj As Object
Dim i, j, tt
Dim node As Object
Dim BasicInfo As Object
Dim Klinfo As Object
Dim Loc As Integer

    Loc = cnt
    Set divObjects = doc.body.All.tags("DIV")
    For i = 0 To divObjects.Length - 1
        Set divObj = divObjects(i)
        If divObj.className = "clearfix container_wrapper pankoudata" Then     '输出的数据
              Set node = divObj.ChildNodes
              Set BasicInfo = node(0).ChildNodes
              Set Klinfo = node(1).Rows
              
              If Loc <= UBound(KLarr) Then    '2018.10.18 add ,由于在第一次取记录后，数据会随时增长，导致读到最后一页时的数据已经发生变更，
                                              '因而将所取数据限定在第一次刷新网页时的记录数。
                  KLarr(Loc, 0) = pDt                 '期数
                  KLarr(Loc, 1) = BasicInfo(0).innerText                   '标识
                  KLarr(Loc, 2) = Right(BasicInfo(0).ChildNodes(0).innerText, 3)    '期数
                  KLarr(Loc, 3) = BasicInfo(0).ChildNodes(1).innerText     '联赛
                  tt = Split(BasicInfo(0).ChildNodes(2).innerText, " ")
                  
                  If DatePart("m", pDt) = 12 And Split(tt(0), "-")(0) = "01" Then
                    KLarr(Loc, 4) = CStr(DatePart("yyyy", pDt) + 1) & "-" & tt(0)  '日期
                  Else
                    KLarr(Loc, 4) = CStr(DatePart("yyyy", pDt)) & "-" & tt(0) '日期
                  End If
                  
                  KLarr(Loc, 5) = tt(1)     '时间
                  For j = 0 To BasicInfo(1).ChildNodes.Length - 1
                      If "SPAN" = BasicInfo(1).ChildNodes(j).nodename Then
                            KLarr(Loc, 6) = BasicInfo(1).ChildNodes(j).innerText     '主队
                      ElseIf "STRONG" = BasicInfo(1).ChildNodes(j).nodename Then
                            KLarr(Loc, 7) = BasicInfo(1).ChildNodes(j).innerText     '比分
                      ElseIf "B" = BasicInfo(1).ChildNodes(j).nodename Then
                            KLarr(Loc, 8) = BasicInfo(1).ChildNodes(j).innerText     '客队
                      End If
                  Next j
                  For j = 0 To Klinfo.Length - 1
                       If InStr(Klinfo(j).innerText, "威廉") > 0 Then
                            KLarr(Loc, 9) = Klinfo(j).Cells(9).innerText                       '主胜
                            KLarr(Loc, 10) = Klinfo(j).Cells(10).innerText                     '平局
                            KLarr(Loc, 11) = Klinfo(j).Cells(11).innerText                     '客胜
                            KLarr(Loc, 12) = Klinfo(j).Cells(8).innerText                     '赔付率
                            
                       End If
                       If InStr(Klinfo(j).innerText, "Bet365") > 0 Then
                            KLarr(Loc, 13) = Klinfo(j).Cells(9).innerText                       '主胜
                            KLarr(Loc, 14) = Klinfo(j).Cells(10).innerText                     '平局
                            KLarr(Loc, 15) = Klinfo(j).Cells(11).innerText                     '客胜
                            KLarr(Loc, 16) = Klinfo(j).Cells(8).innerText                     '赔付率
                       End If
                       If InStr(Klinfo(j).innerText, "澳门") > 0 Then
                            KLarr(Loc, 17) = Klinfo(j).Cells(9).innerText                       '主胜
                            KLarr(Loc, 18) = Klinfo(j).Cells(10).innerText                     '平局
                            KLarr(Loc, 19) = Klinfo(j).Cells(11).innerText                     '客胜
                            KLarr(Loc, 20) = Klinfo(j).Cells(8).innerText                     '赔付率
                       End If
                       If InStr(Klinfo(j).innerText, "所选公司凯利方差") > 0 Then
                            KLarr(Loc, 21) = Klinfo(j).Cells(1).innerText                       '主胜
                            KLarr(Loc, 22) = Klinfo(j).Cells(2).innerText                     '平局
                            KLarr(Loc, 23) = Klinfo(j).Cells(3).innerText                     '客胜
    
                       End If
                  Next
                  Loc = Loc + 1
                End If
        End If
    Next


Set divObjects = Nothing
Set divObj = Nothing
Set node = Nothing
Set BasicInfo = Nothing
Set Klinfo = Nothing

规范澳客网凯利指数 = Loc

End Function








Sub 澳客网盘口评测(begindate As Date, enddate As Date)

Dim doc As Object 'MSHTML.HTMLDocument
Dim txt As String
Dim txt1

Dim recCnt As Integer    '总记录数
Dim pageCnt As Integer   '总页数
Dim k As Integer
Dim i As Integer
Dim j As Integer
Dim t1 As String
Dim t2 As String
Dim tbody As Object
Dim ucell As Object
Dim rowcnt As Integer
Dim totcnt As Integer
Dim data()
Dim WebBrowser1 As Object
Dim wkSheet As Worksheet
Dim dt As Date
Dim sdt As String

'add by ljqu  2018.12.14
Dim URL As String, postData As String
Dim oDoc As Object
'add 2018.12.14 end


Dim waitcnt As Integer   '延时等待次数   ,   2018.11.1  为了处理在加载过程部分页面无响应而导致的程序死循环而设计
Dim errorinfo As String   '记录被装载跳过的页数  2018.11.1  为了处理在加载过程部分页面无响应而导致的程序死循环而设计

Set wkSheet = ActiveWorkbook.Sheets("澳客网(4)")
wkSheet.Cells.Clear

'wkSheet.Cells(1, 1) = "期数"
wkSheet.Cells(1, 2) = "标识"
wkSheet.Cells(1, 3) = "本期编号"
wkSheet.Cells(1, 4) = "联赛"
wkSheet.Cells(1, 5) = "日期"
wkSheet.Cells(1, 6) = "时间"
wkSheet.Cells(1, 7) = "主队"
wkSheet.Cells(1, 8) = "比分"
wkSheet.Cells(1, 9) = "客队"
wkSheet.Cells(1, 10) = "Bet365-初始贴水"
wkSheet.Cells(1, 11) = "盘口"
wkSheet.Cells(1, 12) = "贴水"
wkSheet.Cells(1, 13) = "评测"
wkSheet.Cells(1, 14) = "赛前24小时-贴水"
wkSheet.Cells(1, 15) = "盘口"
wkSheet.Cells(1, 16) = "贴水"
wkSheet.Cells(1, 17) = "评测"
wkSheet.Cells(1, 18) = "赛前8小时-贴水"
wkSheet.Cells(1, 19) = "盘口"
wkSheet.Cells(1, 20) = "贴水"
wkSheet.Cells(1, 21) = "评测"
wkSheet.Cells(1, 22) = "赛前2小时-贴水"
wkSheet.Cells(1, 23) = "盘口"
wkSheet.Cells(1, 24) = "贴水"
wkSheet.Cells(1, 25) = "评测"
wkSheet.Cells(1, 26) = "最新盘口-贴水"
wkSheet.Cells(1, 27) = "盘口"
wkSheet.Cells(1, 28) = "贴水"
wkSheet.Cells(1, 29) = "评测"

'澳门彩票
wkSheet.Cells(1, 30) = "澳门彩票-初始贴水"
wkSheet.Cells(1, 31) = "盘口"
wkSheet.Cells(1, 32) = "贴水"
wkSheet.Cells(1, 33) = "评测"
wkSheet.Cells(1, 34) = "赛前24小时-贴水"
wkSheet.Cells(1, 35) = "盘口"
wkSheet.Cells(1, 36) = "贴水"
wkSheet.Cells(1, 37) = "评测"
wkSheet.Cells(1, 38) = "赛前8小时-贴水"
wkSheet.Cells(1, 39) = "盘口"
wkSheet.Cells(1, 40) = "贴水"
wkSheet.Cells(1, 41) = "评测"
wkSheet.Cells(1, 42) = "赛前2小时-贴水"
wkSheet.Cells(1, 43) = "盘口"
wkSheet.Cells(1, 44) = "贴水"
wkSheet.Cells(1, 45) = "评测"
wkSheet.Cells(1, 46) = "最新盘口-贴水"
wkSheet.Cells(1, 47) = "盘口"
wkSheet.Cells(1, 48) = "贴水"
wkSheet.Cells(1, 49) = "评测"


wkSheet.Columns("C:C").NumberFormatLocal = "@"
wkSheet.Columns("E:E").NumberFormatLocal = "yyyy/m/d"
wkSheet.Columns("H:H").NumberFormatLocal = "@"
totcnt = 1

Set WebBrowser1 = UserForm1.WebBrowser1
dt = begindate
'Do While dt <= enddate
    sdt = CStr(dt)
    sdt = Replace(sdt, "/", "-")
    WebBrowser1.Navigate "http://www.okooo.com/danchang/shuju/pankou/"    '+ sdt
'    WebBrowser1.Refresh
    Do Until WebBrowser1.ReadyState = 4
        DoEvents
    Loop
    Set doc = WebBrowser1.document
    Set tbody = doc.getElementsByTagName("div")

    '***********************************************
    '   由于在实际的页面访问中会出现500错误
    '   因而先将pageCnt和recCnt 置为0
    '   2016.9.4   add by ljqu
    '***********************************************
    recCnt = 0
    pageCnt = 0

    For Each ucell In tbody
        If ucell.className = "pagination" Then    '获取页数信息
            'MsgBox ("页码")
            If ucell.innerText = "暂时没有添加记录" Then   '没有内容，则执行下一天
                recCnt = 0
                pageCnt = 0
                Exit For
            End If
            txt = Split(ucell.innerText, "页首页上一页")(0)
             txt1 = Split(txt, "条记录共")
             pageCnt = txt1(1)
             recCnt = Split(txt1(0), "共有")(1)
             Exit For
        End If
    Next
    
    '处理第一页数据
    If pageCnt > 0 Then
        '标识，期数，联赛，日期，时间，主队，比分，客队，必发赔率(主胜、平局、客胜）、99家平均
        ReDim data(recCnt, 49)
        '球队
        
        postData = "LeagueID="
        Set oDoc = doc.getElementById("csfilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        '让球数据
        postData = postData & "&HandicapNumber="
        Set oDoc = doc.getElementById("rqfilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        
        '数据
        postData = postData & "&BetDate="
        Set oDoc = doc.getElementById("datafilter").getElementsByTagName("input")
        k = 0
        For Each ucell In oDoc
            If ucell.Checked = True Then
                postData = postData & ucell.Value
                k = k + 1
            End If
            If k < oDoc.Length Then
                postData = postData & "%2C"
            End If
        Next
        
        'MakerType
        Set oDoc = doc.getElementById("makerTypeObj")
        postData = postData & "&MakerType=" & oDoc.select_company
        '&  HasEnd
        postData = postData & "&HasEnd=1&PageID="  '2, pageNo
        
    End If

    '处理剩余页面的数据
    errorinfo = ""
    rowcnt = 1
    i = 1
    Do
    'For i = 1 To pageCnt
        Sleep 500 * Round(Rnd, 2) + 500 * Round(Rnd, 2) + 1000 * Round(Rnd, 3)
        
        URL = "http://www.okooo.com/danchang/shuju/pankou?" & postData & i
        Debug.Print URL
        WebBrowser1.Navigate URL   '"javascript:JsGoTo(" + CStr(i) + ")"     '原有分号，2015.7.27去掉“;"
        waitcnt = 0
        Do Until WebBrowser1.ReadyState = 4 Or waitcnt > 30
            Sleep 1000

            waitcnt = waitcnt + 1
            If waitcnt > 30 Then    '超过10S的等待则跳过
                errorinfo = errorinfo & i & ","
            End If
            DoEvents
        Loop
        If WebBrowser1.ReadyState = 4 Then
            Set doc = WebBrowser1.document
            If i = 1 Then   '重新取一次页码数据
                Set tbody = doc.getElementsByTagName("div")
                For Each ucell In tbody
                    If ucell.className = "pagination" Then    '获取页数信息
                        'MsgBox ("页码")
                        If ucell.innerText = "暂时没有添加记录" Then   '没有内容，则执行下一天
                            recCnt = 0
                            pageCnt = 0
                            Exit For
                        End If
                        
                        txt = Split(ucell.innerText, "页首页上一页")(0)
                         txt1 = Split(txt, "条记录共")
                         pageCnt = txt1(1)
                         recCnt = Split(txt1(0), "共有")(1)
                         Exit For
                    End If
                Next
                ReDim data(recCnt, 49)
            End If
            '处理相关资料
            rowcnt = 规范澳客网盘口评测(doc, rowcnt, data, dt)
        End If
    'Next i
    i = i + 1
    Loop While i <= pageCnt
    
    If errorinfo <> "" Then
        MsgBox ("盘口评测数据加载失败的页数包括(" & errorinfo & ")")
    End If
    
    '数据过录
    For i = 1 To recCnt
        For j = 0 To 49
            wkSheet.Cells(i + totcnt, j + 1) = data(i, j)
        Next
    Next
    
    totcnt = totcnt + recCnt
    
    dt = DateAdd("d", 1, dt)   '增加一天
    Sleep 150
'Loop
    wkSheet.Cells(1, 1) = totcnt
    WebBrowser1.Navigate "about:blank"
    Set wkSheet = Nothing
    
End Sub


Function 规范澳客网盘口评测(doc As Object, cnt As Integer, KLarr, pDt As Date)
'------------------------------------------------------------------------------------
'功能说明：  对澳客网盘口评测数据进行规范化
'参数说明：
'         doc： 对应的网页对象
'         cnt：数组开始记录的下标
'         BFarr：记录数据的数组
'         pDt： 澳客网竞彩日期
'-------------------------------------------------------------------------------------
Dim divObjects As Object
Dim divObj As Object
Dim i, j, tt, k
Dim node As Object
Dim BasicInfo As Object
Dim Klinfo As Object
Dim Loc As Integer

    Loc = cnt
    Set divObjects = doc.body.All.tags("DIV")
    For i = 0 To divObjects.Length - 1
        Set divObj = divObjects(i)
        If divObj.className = "clearfix container_wrapper pankoudata" Then     '输出的数据
              Set node = divObj.ChildNodes
              Set BasicInfo = node(0).ChildNodes
              Set Klinfo = node(1).Rows
              
              If Loc <= UBound(KLarr) Then    '2018.10.18 add ,由于在第一次取记录后，数据会随时增长，导致读到最后一页时的数据已经发生变更，
                                              '因而将所取数据限定在第一次刷新网页时的记录数。
                KLarr(Loc, 0) = pDt                 '期数
                KLarr(Loc, 1) = BasicInfo(0).innerText                   '标识
                KLarr(Loc, 2) = Right(BasicInfo(0).ChildNodes(0).innerText, 3)    '期数
                KLarr(Loc, 3) = BasicInfo(0).ChildNodes(1).innerText     '联赛
                tt = Split(BasicInfo(0).ChildNodes(2).innerText, " ")
                If DatePart("m", pDt) = 12 And Split(tt(0), "-")(0) = "01" Then
                  KLarr(Loc, 4) = CStr(DatePart("yyyy", pDt) + 1) & "-" & tt(0)  '日期
                Else
                  KLarr(Loc, 4) = CStr(DatePart("yyyy", pDt)) & "-" & tt(0) '日期
                End If
                KLarr(Loc, 5) = tt(1)     '时间
                For j = 0 To BasicInfo(1).ChildNodes.Length - 1
                    If "SPAN" = BasicInfo(1).ChildNodes(j).nodename Then
                          KLarr(Loc, 6) = BasicInfo(1).ChildNodes(j).innerText     '主队
                    ElseIf "STRONG" = BasicInfo(1).ChildNodes(j).nodename Then
                          KLarr(Loc, 7) = BasicInfo(1).ChildNodes(j).innerText     '比分
                    ElseIf "B" = BasicInfo(1).ChildNodes(j).nodename Then
                          KLarr(Loc, 8) = BasicInfo(1).ChildNodes(j).innerText     '客队
                    End If
                Next j
                For j = 0 To Klinfo.Length - 1
                     If InStr(Klinfo(j).innerText, "Bet365") > 0 Then
                          For k = 1 To Klinfo(j).Cells.Length - 1
                              KLarr(Loc, 8 + k) = Klinfo(j).Cells(k).innerText
                          Next
                     End If
                     If InStr(Klinfo(j).innerText, "澳门") > 0 Then
                          For k = 1 To Klinfo(j).Cells.Length - 1
                              KLarr(Loc, 28 + k) = Klinfo(j).Cells(k).innerText
                          Next
                     End If
                Next
                Loc = Loc + 1
            End If
        End If
    Next


Set divObjects = Nothing
Set divObj = Nothing
Set node = Nothing
Set BasicInfo = Nothing
Set Klinfo = Nothing

规范澳客网盘口评测 = Loc

End Function


