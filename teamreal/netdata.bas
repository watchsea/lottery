Attribute VB_Name = "netdata"
Option Explicit
Sub 球探网赛程积分(teamSeason As String, teamId As String, bgrow As Long)
'TeamSeaSon  赛季
'TeamId      联赛Id
'bgRow       数据写入开始栏
'

Dim doc
Dim tbody As Object
Dim txt As String
Dim txt1
Dim bodyStr As String

Dim loc As Long   '位置


Dim i, j, k, t

Dim wkSheet As Worksheet
Dim WebBrowser1 As Object
Dim winhttp As Object

Dim urlStr As String       '数据IP
Dim mainUrl As String   '原有主页面
Dim cookie

Dim data()
Dim teamData()
Dim arrLeague   '联赛信息
Dim LeagueName As String   '联赛简称
Dim arrTeam     '球队信息
Dim isFinish As Boolean   '场次是否结束

Dim tempDict  As Object   '球队字典
Dim itemId, itemVal

Dim arrCham     '赛事信息

Dim totalRound As Integer  '赛季总轮次
Dim currRound As Integer   '当前进行轮次
Dim roundInfo As String    '轮次信息
Dim roundInfo1 As String
Dim roundInfo2 As Object
Dim roundInfo3, roundInfo4
Dim lunCi  As Integer

Dim myjs


Dim a, b, blen
Dim re As Object

Dim dt, hr



Set tempDict = CreateObject("Scripting.Dictionary")
Set wkSheet = ThisWorkbook.Sheets("赛程积分")

Set re = CreateObject("VBscript.RegExp")
re.Pattern = "&nbsp;\s*"
re.Global = True
re.IgnoreCase = True
're.MultiLine = True

dt = CStr(Year(Now())) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
hr = Left(CStr(Time), 2)


'一级球赛：http://zq.win007.com/cn/League/2014-2015/36.html
'二级球赛：http://zq.win007.com/cn/SubLeague/2011-2012/9.html   ，德乙
'http://zq.win007.com/jsdata/matchResult/2014-2015/s36.js

mainUrl = "http://zq.win007.com/cn/League/" + teamSeason + "/" + teamId + ".html"
urlStr = "http://zq.win007.com/jsData/matchResult/" + teamSeason + "/s" + teamId + ".js" + "?version=" + dt + hr

'urlStr = "http://zq.win007.com/jsData/matchResult/" + teamSeason + "/" + teamId + ".html"
'urlStr = "http://zq.win007.com/jsData/matchResult/2017-2018/s36.js?version=2017091721"
    


     '取网站数据
     Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
     With winhttp
         .Option(6) = 0
         .Open "GET", mainUrl, False              '执行这句，得到的网页数据是英文
         .setRequestHeader "Connection", "Keep-Alive"
         .send
         'cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)    '获取Cookie            '第一次获取Cookie，决定了语言

         'cookie = "Hm_lvt_860f3361e3ed9c994816101d37900758=1408275075,1408362165;" + cookie + ";Hm_lpvt_860f3361e3ed9c994816101d37900758=1408364038"
         .Option(6) = 1
         .Option(2) = 936   '65001      ' 936或950或65001           'GB2312/BIG5/UTF-8
         .Open "GET", urlStr, False  '&_1408364398357", False          '第二次取得数据"
         .setRequestHeader "Referer", mainUrl
         '.setRequestHeader "Cookie", cookie
         .setRequestHeader "Connection", "Keep-Alive"
         .send
    End With
    
    bodyStr = BytesToBstr(winhttp.responsebody, "UTF-8")
    
    '2017.9.22 此处注意，在运行时要注释掉，以免因内存不够而导致失败
    'With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'DataObject对象，数据放入剪贴板，记事本观察数据10. ? ?? ?? ?? ?
    '        .SetText bodyStr  '因为XMLHTTP默认是UTF-8，不能识别gb2312，会发现数据乱码
    '        .PutInClipboard   '所以不能采用.responsetext对象来得到字符串
    'End With

    'WebBrowser1.Navigate "about:blank"
    'data = Filter(Filter(Split(Trim(winhttp.responsetext), vbCrLf), "var", True), "!", False)
    Set myjs = CreateObject("MSScriptControl.ScriptControl") '调用ScriptControl对象将提取的变量文本运算形成对象集合

   myjs.Language = "javascript"

   
    If InStr(bodyStr, "查看的页面不存在") > 0 Then
        Exit Sub
    End If
 
    txt = re.Replace(bodyStr, "")
    txt1 = Split(txt, ";")
    For i = 0 To UBound(txt1)
        If InStr(txt1(i), "var") > 0 Then
            myjs.addcode (txt1(i)) '输入
        End If
    Next
    Set a = myjs.CodeObject.arrLeague '联赛信息
    Set b = myjs.CodeObject.arrTeam   '球队信息
    
    '联赛信息
    arrLeague = Split(a, ",")

    '联赛轮次
    totalRound = arrLeague(7)
    LeagueName = arrLeague(9)
    currRound = arrLeague(8)
    
        
    '球队信息
    blen = CallByName(b, "length", VbGet)
    ReDim teamData(blen, 5)
    
    loc = 1
    For Each arrTeam In b
        itemId = CallByName(arrTeam, "0", VbGet)
        itemVal = CallByName(arrTeam, "1", VbGet)
        If tempDict.exists(itemId) Then
            tempDict.Item(itemId) = itemVal
        Else
            tempDict.Add itemId, itemVal
        End If
        teamData(loc, 1) = teamId     '联赛Id
        teamData(loc, 2) = LeagueName  '联赛名称
        teamData(loc, 3) = itemId     '球队Id
        teamData(loc, 4) = itemVal    '球队名称
        teamData(loc, 5) = teamSeason   '赛季
        loc = loc + 1
    Next
    
    '此次原为totalRound*10，由于美职业每轮次不只10场比赛，因而改为
    '1000场，每个赛季每个联赛的场次不可能超过1000次
    ReDim data(1000, 27) '轮次，场次（每轮10场），每场23个数据项，后面加联赛名称，主队名称，客队名称
    
    '处理每轮赛事信息
    loc = 0
    i = 1
    isFinish = False

    While i <= totalRound And Not isFinish
    'For i = 1 To totalRound
        roundInfo = "jh[""R_" & i & """] = "
        For j = 0 To UBound(txt1)
            If InStr(txt1(j), roundInfo) > 0 Then
                roundInfo1 = Replace(txt1(j), roundInfo, "")
                '处理数据
                Call getItemfromJson(roundInfo1, roundInfo2)
                lunCi = UBound(Split(roundInfo1, "],["))
                For k = 0 To lunCi
                    Set roundInfo3 = CallByName(roundInfo2, k, VbGet)
                    roundInfo4 = CallByName(roundInfo3, 6, VbGet)
                    
                    If i <= currRound Then             '如果轮次或者等于当前轮次，则不判断是否有积分，2016.8.12加上
                        data(loc, 1) = i    '轮次
                        data(loc, 0) = teamSeason  '赛季
                        For t = 0 To 22
                            On Error Resume Next
                            data(loc, t + 2) = CallByName(roundInfo3, t, VbGet) '(t)
                        Next t
                        data(loc, 25) = LeagueName                   '联赛名称
                        data(loc, 26) = tempDict.Item(data(loc, 6))   '主队名称
                        data(loc, 27) = tempDict.Item(data(loc, 7))   '客队名称
                        loc = loc + 1
                        
                    Else          '超过当前轮次，则只抓取已打完的赛事
                        If InStr(roundInfo4, "-") > 0 Then '是正常的得分数据，
                            data(loc, 1) = i    '轮次
                            data(loc, 0) = teamSeason  '赛季
                            For t = 0 To 22
                                On Error Resume Next
                                data(loc, t + 2) = CallByName(roundInfo3, t, VbGet) '(t)
                            Next t
                            data(loc, 25) = LeagueName                   '联赛名称
                            data(loc, 26) = tempDict.Item(data(loc, 6))   '主队名称
                            data(loc, 27) = tempDict.Item(data(loc, 7))   '客队名称
                            loc = loc + 1
                        End If
                    End If
                Next k
                '数据登记完毕
                Exit For
            End If
        Next
    'Next
    i = i + 1
    Wend

    
    For i = 0 To loc - 1
        For j = 0 To 27
            wkSheet.Cells(bgrow + i, j + 1) = CStr(data(i, j))
        Next j
    Next
    
    '更新球队信息
    Call 更新球队信息(teamData)
    
    bgrow = bgrow + loc
    Set wkSheet = Nothing
    
    
End Sub


Sub 网站数据更新()
Dim wkSheet As Worksheet
Dim leagueSheet As Worksheet    '联赛名册
Dim bgrow As Long
Dim cnt As Long
Dim teamSeason As String     '赛季
Dim seasonType               '赛季类型
Dim teamType             '球赛类型
Dim teamId As String        '联赛ID
Dim subTeamId As String     '二级联赛Id



Dim i As Long

Call 初始化字典(Dict, "Param")

Set wkSheet = ThisWorkbook.Sheets("赛程积分")
Set leagueSheet = ThisWorkbook.Sheets("LeagueConfig")

wkSheet.Cells.Clear

wkSheet.Columns("A:A").NumberFormatLocal = "@"
wkSheet.Columns("I:I").NumberFormatLocal = "@"
'wkSheet.Columns("F:F").NumberFormatLocal = "yyyy-mm-dd  HH:MM:SS"
wkSheet.Columns("J:J").NumberFormatLocal = "@"

wkSheet.Cells(1, 1) = "赛季"
wkSheet.Cells(1, 2) = "轮次"
wkSheet.Cells(1, 3) = "球赛Id"
wkSheet.Cells(1, 4) = "联赛Id"
wkSheet.Cells(1, 5) = ""
wkSheet.Cells(1, 6) = "日期"
wkSheet.Cells(1, 7) = "主队ID"
wkSheet.Cells(1, 8) = "客队ID"
wkSheet.Cells(1, 9) = "比分"
wkSheet.Cells(1, 10) = "半场比分"
wkSheet.Cells(1, 11) = "主队积分排名"
wkSheet.Cells(1, 12) = "客队积分排名"
wkSheet.Cells(1, 13) = "全球让球"
wkSheet.Cells(1, 14) = "半球让球"
wkSheet.Cells(1, 15) = "全场大小"
wkSheet.Cells(1, 16) = "半场大小"
wkSheet.Cells(1, 17) = "析"
wkSheet.Cells(1, 18) = "欧"
wkSheet.Cells(1, 19) = "亚"
wkSheet.Cells(1, 20) = "大"
wkSheet.Cells(1, 21) = "主队红牌次数"
wkSheet.Cells(1, 22) = "客队红牌次数"
wkSheet.Cells(1, 26) = "联赛简称"
wkSheet.Cells(1, 27) = "主队名称"
wkSheet.Cells(1, 28) = "客队名称"


bgrow = 2

cnt = leagueSheet.UsedRange.Rows(leagueSheet.UsedRange.Rows.Count).row
For i = 2 To cnt
    teamId = leagueSheet.Cells(i, 1)
    seasonType = leagueSheet.Cells(i, 4)
    teamType = leagueSheet.Cells(i, 6)
    subTeamId = leagueSheet.Cells(i, 7)
    '处理赛季信息：1：表示跨年赛季，如2015-2016，2：表示非跨年赛季：如2015
    
    If seasonType = 2 Then     '如果赛季为非跨年寒季，则取前四位
        teamSeason = Dict.Item("TEAMSEASON2")    'Left(teamSeason, 4)
    Else
        teamSeason = Dict.Item("TEAMSEASON1")
    End If
    
    If teamType = 2 Then    '二级联赛
        'teamId = teamId & "_" & subTeamId
        Call 球探网二级联赛赛程积分(teamSeason, teamId, subTeamId, bgrow)
    Else
        Call 球探网赛程积分(teamSeason, teamId, bgrow)
    End If

    'MsgBox (bgRow)
Next

MsgBox ("网站数据处理完毕!")

End Sub


Sub 更新球队信息(teamData)
'teamData   球队数据

Dim wkSheet As Worksheet
Dim row1  As Long
Dim teamDict As Object     '球队字典
Dim i, j
Dim id

Call 初始化球队字典(teamDict, "球队信息")

Set wkSheet = ThisWorkbook.Sheets("球队信息")
row1 = wkSheet.UsedRange.Rows(wkSheet.UsedRange.Rows.Count).row

For i = 1 To UBound(teamData)
    If teamData(i, 1) <> "" Then
        id = teamData(i, 1) & teamData(i, 3) & teamData(i, 5)
        If Not teamDict.exists(id) Then    '不存在则添加
            row1 = row1 + 1
            For j = 1 To 5
                wkSheet.Cells(row1, j) = teamData(i, j)
            Next
        End If
    End If
Next

Set wkSheet = Nothing
Set teamDict = Nothing

End Sub



Sub 球探网二级联赛赛程积分(teamSeason As String, teamId As String, subTeamId As String, bgrow As Long)
'TeamSeaSon  赛季
'TeamId      联赛Id
'bgRow       数据写入开始栏
'

Dim doc
Dim tbody As Object
Dim txt As String
Dim txt1
Dim bodyStr As String

Dim loc As Long   '位置


Dim i, j, k, t

Dim wkSheet As Worksheet
Dim WebBrowser1 As Object
Dim winhttp As Object

Dim urlStr As String    '数据IP
Dim mainUrl As String   '原有主页面
Dim cookie

Dim data()
Dim teamData()
Dim arrLeague   '联赛信息
Dim LeagueName As String   '联赛简称
Dim arrTeam     '球队信息
Dim arrSubLeague  '二级联赛信息
Dim isFinish As Boolean   '场次是否结束

Dim tempDict  As Object   '球队字典
Dim itemId, itemVal

Dim arrCham     '赛事信息

Dim totalRound As Integer   '赛季总轮次
Dim currRound As Integer   '当前进行轮次
Dim roundInfo As String    '轮次信息
Dim roundInfo1 As String
Dim roundInfo2 As Object
Dim roundInfo3, roundInfo4
Dim lunCi  As Integer

Dim myjs


Dim a, b, c, blen
Dim re As Object

Dim dt, hr


Set tempDict = CreateObject("Scripting.Dictionary")
Set wkSheet = ThisWorkbook.Sheets("赛程积分")

Set re = CreateObject("VBscript.RegExp")
re.Pattern = "&nbsp;\s*"
re.Global = True
re.IgnoreCase = True
're.MultiLine = True

dt = CStr(Year(Now())) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)

hr = Left(CStr(Time), 2)

Set WebBrowser1 = UserForm1.WebBrowser1
'一级球赛：http://zq.win007.com/cn/League/2014-2015/36.html
'二级球赛：http://zq.win007.com/cn/SubLeague/2011-2012/9.html   ，德乙
'http://zq.win007.com/jsdata/matchResult/2014-2015/s36.js

mainUrl = "http://zq.win007.com/cn/SubLeague/" + teamSeason + "/" + teamId + ".html"
urlStr = "http://zq.win007.com/jsData/matchResult/" + teamSeason + "/s" + teamId + "_" + subTeamId + ".js" + "?version=" + dt + hr


    '取网站数据
     Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
     With winhttp
         .Option(6) = 0
         .Open "GET", mainUrl, False              '执行这句，得到的网页数据是英文
         .setRequestHeader "Connection", "Keep-Alive"
         .send
         'cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)    '获取Cookie            '第一次获取Cookie，决定了语言

         'cookie = "Hm_lvt_860f3361e3ed9c994816101d37900758=1408275075,1408362165;" + cookie + ";Hm_lpvt_860f3361e3ed9c994816101d37900758=1408364038"
         .Option(6) = 1
         .Option(2) = 936   '65001      ' 936或950或65001           'GB2312/BIG5/UTF-8
         .Open "GET", urlStr, False  '&_1408364398357", False          '第二次取得数据"
         .setRequestHeader "Referer", mainUrl
         '.setRequestHeader "Cookie", cookie
         .setRequestHeader "Connection", "Keep-Alive"
         .send
    End With
    
    bodyStr = BytesToBstr(winhttp.responsebody, "UTF-8")
    
    '2017.9.22 此处注意，在运行时要注释掉，以免因内存不够而导致失败
    'With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'DataObject对象，数据放入剪贴板，记事本观察数据10. ? ?? ?? ?? ?
    '        .SetText bodyStr  '因为XMLHTTP默认是UTF-8，不能识别gb2312，会发现数据乱码
    '        .PutInClipboard   '所以不能采用.responsetext对象来得到字符串
    'End With



    'WebBrowser1.Navigate "about:blank"
    'data = Filter(Filter(Split(Trim(winhttp.responsetext), vbCrLf), "var", True), "!", False)
    Set myjs = CreateObject("MSScriptControl.ScriptControl") '调用ScriptControl对象将提取的变量文本运算形成对象集合

   myjs.Language = "javascript"

   
    If InStr(bodyStr, "查看的页面不存在") > 0 Then
        Exit Sub
    End If
 
    txt = re.Replace(bodyStr, "")
    txt1 = Split(txt, ";")
    For i = 0 To UBound(txt1)
        If InStr(txt1(i), "var") > 0 Then
            myjs.addcode (txt1(i)) '输入
        End If
    Next
    Set a = myjs.CodeObject.arrLeague '联赛信息
    Set b = myjs.CodeObject.arrTeam   '球队信息
    Set c = myjs.CodeObject.arrSubLeague '二级联赛信息
    
    '联赛信息
    arrLeague = Split(a, ",")
    
    arrSubLeague = Split(c, ",")

    '联赛轮次
    totalRound = arrSubLeague(5)    '总轮次
    currRound = arrSubLeague(6)     '当前进行轮次
    LeagueName = arrLeague(8)
    
         
    '球队信息
    blen = CallByName(b, "length", VbGet)
    ReDim teamData(blen, 5)
    
    loc = 1
    For Each arrTeam In b
        itemId = CallByName(arrTeam, "0", VbGet)
        itemVal = CallByName(arrTeam, "1", VbGet)
        If tempDict.exists(itemId) Then
            tempDict.Item(itemId) = itemVal
        Else
            tempDict.Add itemId, itemVal
        End If
        teamData(loc, 1) = teamId     '联赛Id
        teamData(loc, 2) = LeagueName  '联赛名称
        teamData(loc, 3) = itemId     '球队Id
        teamData(loc, 4) = itemVal    '球队名称
        teamData(loc, 5) = teamSeason   '赛季
        loc = loc + 1
    Next
         
    
    '此次原为totalRound*10，由于美职业每轮次不只10场比赛，因而改为
    '1000场，每个赛季每个联赛的场次不可能超过1000次
    ReDim data(1000, 27) '轮次，场次（每轮10场），每场23个数据项，后面加联赛名称，主队名称，客队名称
    
    '处理每轮赛事信息
    loc = 0
    i = 1
    isFinish = False

    While i <= totalRound And Not isFinish
    'For i = 1 To totalRound
        roundInfo = "jh[""R_" & i & """] = "
        For j = 0 To UBound(txt1)
            If InStr(txt1(j), roundInfo) > 0 Then
                roundInfo1 = Replace(txt1(j), roundInfo, "")
                '处理数据
                Call getItemfromJson(roundInfo1, roundInfo2)
                lunCi = UBound(Split(roundInfo1, "],["))
                If lunCi > 0 Then
                    For k = 0 To lunCi
                        Set roundInfo3 = CallByName(roundInfo2, k, VbGet)
                        roundInfo4 = CallByName(roundInfo3, 6, VbGet)
                        
                        '没有得分信息，则表明还没有比赛，后继的赛事也没有进行，2015.9.24 delete
                        'If roundInfo4 = "" Then
                            'isFinish = True
                            'Exit For
                        'End If
                        If i <= currRound Then             '如果轮次或者等于当前轮次，则不判断是否有积分，2016.8.12加上
                            data(loc, 1) = i    '轮次
                            data(loc, 0) = teamSeason  '赛季
                            For t = 0 To 22
                                On Error Resume Next
                                data(loc, t + 2) = CallByName(roundInfo3, t, VbGet) '(t)
                            Next t
                            data(loc, 25) = LeagueName                   '联赛名称
                            data(loc, 26) = tempDict.Item(data(loc, 6))   '主队名称
                            data(loc, 27) = tempDict.Item(data(loc, 7))   '客队名称
                            loc = loc + 1
                        Else
                            If InStr(roundInfo4, "-") > 0 Then '是正常的得分数据，
                                data(loc, 1) = i    '轮次
                                data(loc, 0) = teamSeason  '赛季
                                For t = 0 To 22
                                    On Error Resume Next
                                    data(loc, t + 2) = CallByName(roundInfo3, t, VbGet) '(t)
                                Next t
                                data(loc, 25) = LeagueName                   '联赛名称
                                data(loc, 26) = tempDict.Item(data(loc, 6))   '主队名称
                                data(loc, 27) = tempDict.Item(data(loc, 7))   '客队名称
                                loc = loc + 1
                            End If
                        End If
                    Next k
                End If
                '数据登记完毕
                Exit For
            End If
        Next
    'Next
    i = i + 1
    Wend

    
    For i = 0 To loc - 1
        For j = 0 To 27
            wkSheet.Cells(bgrow + i, j + 1) = CStr(data(i, j))
        Next j
    Next
    
    '更新球队信息
    Call 更新球队信息(teamData)
    
    bgrow = bgrow + loc
    Set wkSheet = Nothing
    
    
End Sub


