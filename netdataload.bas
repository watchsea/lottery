Attribute VB_Name = "netdataload"
Option Explicit
#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Sub 网站数据更新()
Dim matchdate As Date
Dim begindate As Date
Dim enddate As Date

'状态信息
Dim infoStr As String
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim s8 As String
Dim s9 As String



'记录运行时间
Dim bgTm

Dim exectm

s1 = "1.球探网威廉希尔数据"
s2 = "2.球探网BF数据"
s3 = "3.球探网赛事积分数据"
s8 = "4.竞彩网数据"


matchdate = Date
    
    Call 初始化字典(leagueDict, "01赛事")
    
    bgTm = Now
    Application.StatusBar = "球探网威廉希尔数据导入【" & bgTm & "】......"
    Call 球探网数据载入("球探网(W)", "id=115&company=威廉希尔(英国)", matchdate)
    exectm = DateDiff("m", bgTm, Now)
    s1 = s1 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s1 & Chr(10) & s2 & Chr(10) & s3 & Chr(10) & s8
    MsgBox (infoStr)
    
    
    bgTm = Now
    Application.StatusBar = "球探网BF数据导入【" & bgTm & "】......"
    Call 球探网BF数据载入
    exectm = DateDiff("m", bgTm, Now)
    s2 = s2 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s1 & Chr(10) & s2 & Chr(10) & s3 & Chr(10) & s8
    MsgBox (infoStr)
    
    
    
    bgTm = Now
    Application.StatusBar = "球探网赛事积分数据导入【" & bgTm & "】......"
    Call 球探网赛事积分数据载入
    exectm = DateDiff("n", bgTm, Now)
    s3 = s3 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s1 & Chr(10) & s2 & Chr(10) & s3 & Chr(10) & s8
    MsgBox (infoStr)
    
    
    
    
    bgTm = Now
    Application.StatusBar = "竞彩网数据导入【" & bgTm & "】......"
    Call 中国竞彩网数据载入
    'add 2016.8.11 by ljqu
    Application.StatusBar = "竞彩网投资比例数据导入【" & bgTm & "】......"
    Call 竞彩网投资比例数据载入    '此处有依赖关系，依赖于前一个数据的导入
    Application.StatusBar = "竞彩网比分数据导入【" & bgTm & "】......"
    Call 竞彩网比分数据载入
    Application.StatusBar = "竞彩网总进球数据导入【" & bgTm & "】......"
    Call 竞彩网总进球数据载入
    Application.StatusBar = "竞彩网半全场胜平负数据导入【" & bgTm & "】......"
    Call 竞彩网半全场胜平负数据载入
    exectm = DateDiff("n", bgTm, Now)
    s8 = s8 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s1 & Chr(10) & s2 & Chr(10) & s3 & Chr(10) & s8
    MsgBox (infoStr)
    
    Application.StatusBar = "球探网和竞彩网站数据导入完毕！"
    MsgBox ("球探网和竞彩网站数据导入完毕！")
    
    
End Sub


Sub 澳客网数据更新()
Dim matchdate As Date
Dim begindate As Date
Dim enddate As Date

'状态信息
Dim infoStr As String
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim s4 As String
Dim s5 As String
Dim s6 As String
Dim s7 As String
Dim s8 As String
Dim s9 As String



'记录运行时间
Dim bgTm

Dim exectm

s4 = "1.澳客网必发盈亏数据"
s5 = "2.澳客网胜负指数数据"
s6 = "3.澳客网盘口评测数据"
s7 = "4.澳客网凯利指数数据"
s9 = "5.澳客网期数数据"


matchdate = Date
    
    Call 初始化字典(leagueDict, "01赛事")
    
    '澳客网数据导入
    
    bgTm = Now
    begindate = DateAdd("d", -1, matchdate)
    enddate = DateAdd("d", 2, matchdate)
    Application.StatusBar = "澳客网必发盈亏数据导入【" & bgTm & "】......"
    Call 澳客网必发盈亏(begindate, enddate)
    exectm = DateDiff("n", bgTm, Now)
    s4 = s4 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s4 & Chr(10) & s5 & Chr(10) & s6 & Chr(10) & s7 & Chr(10) & s9
    MsgBox (infoStr)
    Sleep (2000 * (Rnd(2) + 1))
    
    bgTm = Now
    Application.StatusBar = "澳客网胜负指数数据导入【" & bgTm & "】......"
    Call 澳客网胜负指数(begindate, enddate)
    exectm = DateDiff("n", bgTm, Now)
    s5 = s5 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s4 & Chr(10) & s5 & Chr(10) & s6 & Chr(10) & s7 & Chr(10) & s9
    MsgBox (infoStr)
    Sleep (2000 * (Rnd(2) + 1))
    
    bgTm = Now
    Application.StatusBar = "澳客网盘口评测数据导入【" & bgTm & "】......"
    Call 澳客网盘口评测(begindate, enddate)
    exectm = DateDiff("n", bgTm, Now)
    s6 = s6 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s4 & Chr(10) & s5 & Chr(10) & s6 & Chr(10) & s7 & Chr(10) & s9
    MsgBox (infoStr)
    Sleep (2000 * (Rnd(2) + 1))
    
    bgTm = Now
    Application.StatusBar = "澳客网凯利指数数据导入【" & bgTm & "】......"
    Call 澳客网凯利指数(begindate, enddate)
    exectm = DateDiff("n", bgTm, Now)
    s7 = s7 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s4 & Chr(10) & s5 & Chr(10) & s6 & Chr(10) & s7 & Chr(10) & s9
    MsgBox (infoStr)
    Sleep (2000 * (Rnd(2) + 1))
    
    bgTm = Now
    Application.StatusBar = "澳客网数据导入【" & bgTm & "】......"
    Call 澳客网数据载入
    exectm = DateDiff("n", bgTm, Now)
    s9 = s9 & "，已完成导入，耗时：" & exectm & "分"
    infoStr = s4 & Chr(10) & s5 & Chr(10) & s6 & Chr(10) & s7 & Chr(10) & s9
    MsgBox (infoStr)
    
    Application.StatusBar = "澳客网数据导入完毕！"
    MsgBox ("澳客网数据导入完毕！")
    
    
End Sub





Sub 球探网BF数据载入()
'------------------------------------------------------------
'dataBF:数据输出的数组
'dataW:要查找的数据根据data(,0)的ID号去链接新的网址数据
'------------------------------------------------------------
Dim rowNo
Dim col
Dim i, j
Dim vsId As String
Dim data()
Dim bfData() As Double
Dim srcdata()
Dim wkSheet As Worksheet


Call LoadDataToArray(srcdata, "球探网(W)")

rowNo = UBound(srcdata, 1)
ReDim data(rowNo, 64)     'id号，四个指标（赔1、赔2、bf1、bf3），前2个指标6个数据，后2个指标各有8个
    '1-4  :赔1：初始值（胜平负）  5-8  :赔1：即时值（胜平负）
    '9-12：赔2：初始值（胜平负）  13-16：赔2：即时值（胜平负）
    '17-20：bf1:初始值（胜、平、负、返还率）  21——24：bf1:即时值（胜、平、负、返还率）
    '25-28：bf2:初始值（胜、平、负、返还率）  29——32：bf2:即时值（胜、平、负、返还率）
    
    
    '33-36：Bet365:初始值（胜、平、负、返还率）         37——40：Bet365:即时值（胜、平、负、返还率）
    '41-44：澳门:初始值（胜、平、负、返还率）           45——48：澳门:即时值（胜、平、负、返还率）
    '49-52：立博（英国）:初始值（胜、平、负、返还率）   53——56：立博（英国）:即时值（胜、平、负、返还率）
    '57-60：易胜博:初始值（胜、平、负、返还率）         61——64：易胜博:即时值（胜、平、负、返还率）
For i = 1 To rowNo
    vsId = srcdata(i, 0)
    Sleep 150
    '根据vsId获取相应的数据组，四个指标数据
    'Call 取欧赔指数(bfData, CStr(vsId))
    
    If 取欧赔指数(bfData, CStr(vsId)) Then     '如果有数据
    '拼装成新形式的数据格式
        data(i, 0) = vsId
        '赔1：1-4,5-8
        col = 1
        For j = 0 To 3    '7-10,
            data(i, col + j) = bfData(1, j + 7) / 100
            data(i, col + j + 4) = bfData(1, j + 7) / 100
        Next
        
        '赔2：9-12，13-16
        col = 9
        For j = 0 To 3
            data(i, col + j) = bfData(1, j + 14) / 100
            data(i, col + j + 4) = bfData(1, j + 14) / 100
        Next
        
        'Beffair（英国）,bf1  17-20,21-24
        col = 17
        For j = 0 To 3
            data(i, col + j) = bfData(0, j + 7) / 100
            data(i, col + j + 4) = bfData(0, j + 14) / 100
        Next
        
        'Bf2,25-28,29-32
        col = 25
        For j = 0 To 2
            data(i, col + j) = bfData(0, j + 18)
            data(i, col + j + 4) = bfData(0, j + 18)
        Next
        
        'Bet365（英国）,bf1  17-20,21-24
        col = 33
        For j = 0 To 3
            data(i, col + j) = bfData(2, j + 7) / 100
            data(i, col + j + 4) = bfData(2, j + 14) / 100
        Next
        
        
        '澳门  17-20,21-24
        col = 41
        For j = 0 To 3
            data(i, col + j) = bfData(3, j + 7) / 100
            data(i, col + j + 4) = bfData(3, j + 14) / 100
        Next
        
        
        '立博（英国）  17-20,21-24
        col = 49
        For j = 0 To 3
            data(i, col + j) = bfData(4, j + 7) / 100
            data(i, col + j + 4) = bfData(4, j + 14) / 100
        Next
        
        
        '易胜博（安提瓜和巴布达）,bf1  17-20,21-24
        col = 57
        For j = 0 To 3
            data(i, col + j) = bfData(5, j + 7) / 100
            data(i, col + j + 4) = bfData(5, j + 14) / 100
        Next
        
    End If
Next

Set wkSheet = ActiveWorkbook.Sheets("球探网(BF)")
wkSheet.Cells.ClearContents

wkSheet.Cells(1, 1) = "序号"
wkSheet.Cells(1, 2) = "赔1-胜"
wkSheet.Cells(1, 3) = "平"
wkSheet.Cells(1, 4) = "负"
wkSheet.Cells(1, 5) = "返还率"

wkSheet.Cells(1, 10) = "赔2-胜"
wkSheet.Cells(1, 11) = "平"
wkSheet.Cells(1, 12) = "负"
wkSheet.Cells(1, 13) = "返还率"

wkSheet.Cells(1, 18) = "Beffair_胜"
wkSheet.Cells(1, 19) = "平"
wkSheet.Cells(1, 20) = "负"
wkSheet.Cells(1, 21) = "返还率"
wkSheet.Cells(1, 22) = "胜2"
wkSheet.Cells(1, 23) = "平2"
wkSheet.Cells(1, 24) = "负2"
wkSheet.Cells(1, 25) = "返还率2"

wkSheet.Cells(1, 26) = "凯利-胜"
wkSheet.Cells(1, 27) = "平"
wkSheet.Cells(1, 28) = "负"


wkSheet.Cells(1, 34) = "Bet365_胜"
wkSheet.Cells(1, 35) = "平"
wkSheet.Cells(1, 36) = "负"
wkSheet.Cells(1, 37) = "返还率"
wkSheet.Cells(1, 38) = "胜2"
wkSheet.Cells(1, 39) = "平2"
wkSheet.Cells(1, 40) = "负2"
wkSheet.Cells(1, 41) = "返还率2"




wkSheet.Cells(1, 42) = "澳门_胜"
wkSheet.Cells(1, 43) = "平"
wkSheet.Cells(1, 44) = "负"
wkSheet.Cells(1, 45) = "返还率"
wkSheet.Cells(1, 46) = "胜2"
wkSheet.Cells(1, 47) = "平2"
wkSheet.Cells(1, 48) = "负2"
wkSheet.Cells(1, 49) = "返还率2"




wkSheet.Cells(1, 50) = "立博_胜"
wkSheet.Cells(1, 51) = "平"
wkSheet.Cells(1, 52) = "负"
wkSheet.Cells(1, 53) = "返还率"
wkSheet.Cells(1, 54) = "胜2"
wkSheet.Cells(1, 55) = "平2"
wkSheet.Cells(1, 56) = "负2"
wkSheet.Cells(1, 57) = "返还率2"



wkSheet.Cells(1, 58) = "易胜博_胜"
wkSheet.Cells(1, 59) = "平"
wkSheet.Cells(1, 60) = "负"
wkSheet.Cells(1, 61) = "返还率"
wkSheet.Cells(1, 62) = "胜2"
wkSheet.Cells(1, 63) = "平2"
wkSheet.Cells(1, 64) = "负2"
wkSheet.Cells(1, 65) = "返还率2"



For i = 1 To rowNo
    For j = 0 To 64
        wkSheet.Cells(i + 1, j + 1) = data(i, j)
    Next
Next

End Sub





Function 取欧赔指数(dataAvg, ids As String)
'------------------------------------------------------------------
'dataAvg:返回的平均值及Beffair值
'ids:球赛对应的id号
'------------------------------------------------------------------

Dim k As Integer
Dim i As Integer
Dim j As Integer
Dim rowcnt As Integer
Dim colCnt As Integer
Dim bfrow  As Integer    'Beffair公司的数据所在行
Dim brow As Integer      'bet365 数据所在行
Dim mrow As Integer      '澳门数据所在行
Dim lrow As Integer      '立博数据所在行
Dim erow As Integer      '易胜博数据所在行
Dim data1()
Dim dataSum()  As Double '汇总值
Dim company, companyinfo
Dim cols
Dim index       'game=Array的在文本中的索引位置

Dim winhttp As Object
Dim tt
Dim tt1, tt2, tt3
Dim URL

'url = "http://1x2.nowscore.com/" + ids + ".js"
'2018.3.4 网站变更取数方式
URL = "http://1x2d.win007.com/" + ids + ".js"
        Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
     With winhttp
         .Option(6) = 1
         .Option(2) = 936   '65001      ' 936或950或65001           'GB2312/BIG5/UTF-8
         .Open "GET", URL, False         '第二次取得数据"
         .setRequestHeader "Connection", "Keep-Alive"
         .send
         '将二进制转换UTF-8
         tt = BytesToBstr(.responseBody, "UTF-8")
         
     End With
     
     index = InStr(tt, "game=Array(")
     If index > 0 Then     '找到game=Array(字串，表明有欧赔数据
         tt1 = Mid(tt, index + 12, Len(tt))
         tt3 = Split(tt1, ";")(0)
         'tt3 = Mid(tt2, 7, Len(tt2) - 7)
         
         'tt3 = tt2
    
         '置换掉"
         
         company = Split(tt3, """,""")
         rowcnt = UBound(company)
         cols = Split(company(0), "|")     '此处原为1，当数据只有一个时，出现下标越界错误，因而改为0，20150718
         colCnt = UBound(cols)
         
         ReDim data1(rowcnt + 1, colCnt + 1)
         ReDim dataSum(colCnt + 1)
         ReDim dataAvg(5, colCnt + 1)
         For i = 0 To rowcnt
            companyinfo = company(i)
            companyinfo = Replace(companyinfo, """", "")     '各个公司的预测数据
            cols = Split(companyinfo, "|")                  '各个公司的预测数据枚举
            Select Case cols(0)
                Case 2:  bfrow = i + 1       '取beffair（英国）数据
                Case 281: brow = i + 1       'Bet365 数据
                Case 80:  mrow = i + 1       '澳门数据
                Case 82:  lrow = i + 1       '立博数据
                Case 90:  erow = i + 1       '易胜博数据
            End Select
            For j = 0 To colCnt
                If j <= 9 Then
                    data1(i + 1, j + 1) = cols(j)            '取初时值
                ElseIf cols(j) = "" And j <= 16 Then         '如果即时值没有，则填入初始值
                    data1(i + 1, j + 1) = cols(j - 7)
                Else
                    data1(i + 1, j + 1) = cols(j)            '有即时值，则填入即时值
                End If
            Next
            '对于每行数据进行加总
            For k = 4 To 20
                    If IsNumeric(data1(i + 1, k)) Then
                        dataSum(k) = dataSum(k) + CDbl(data1(i + 1, k))
                    End If
            Next
         Next
         
         '求平均值及平移beffair指标数据
         For k = 4 To 20
            dataAvg(1, k) = Round(dataSum(k) / (rowcnt + 1), 2)
            '存储数据--BetFair
            If bfrow > 0 And IsNumeric(data1(bfrow, k)) Then
                dataAvg(0, k) = data1(bfrow, k)
            End If
            
            '存储数据--Bet365
            If brow > 0 And IsNumeric(data1(brow, k)) Then
                dataAvg(2, k) = data1(brow, k)
            End If
            
            '存储数据--澳门
            If mrow > 0 And IsNumeric(data1(mrow, k)) Then
                dataAvg(3, k) = data1(mrow, k)
            End If
            
            '存储数据--立博
            If lrow > 0 And IsNumeric(data1(lrow, k)) Then
                dataAvg(4, k) = data1(lrow, k)
            End If
            
            '存储数据--易胜博
            If erow > 0 And IsNumeric(data1(erow, k)) Then
                dataAvg(5, k) = data1(erow, k)
            End If
            
         Next
         取欧赔指数 = True
     Else
        取欧赔指数 = False
     End If
     Set winhttp = Nothing

End Function



Sub 球探网数据载入(sheetName As String, ids, matchdate As Date)
'------------------------------------------------------------------
'ids：相关公司对应的进入参数
'matchdate：历史数据的日期
'dataType:取数据的类型：C-默认值，对当天的数据，H-取历史数据
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



'ids = "1014941"   '"987108"

URL = "http://op1.win007.com/company.aspx?"                  '没有加参数type=1的表示当前以后所有数据

URL = URL + ids

Set IE = UserForm1.WebBrowser1

With IE
  .Navigate URL '网址
  Do Until .ReadyState = 4
    DoEvents
  Loop
  Set doc = .document
End With
'Application.ScreenUpdating = False

Set tt = doc.getElementById("table_schedule").getElementsbyTagName("tr")
rowcnt = tt.Length - 1
colCnt = tt(0).Cells.Length - 1
col = 0
ReDim data1(rowcnt, colCnt + 7)  '重置数组

'读取表头
For j = 0 To tt(0).Cells.Length - 1
    data1(col, j) = tt(0).Cells(j).innerText
Next
For k = 4 To 10
    data1(col, k - 4 + j) = tt(0).Cells(k).innerText
Next

For i = 1 To rowcnt    '处理数据
    tt2 = tt(i).getAttribute("name")   '取名称，即联赛的ID
    If Not IsNull(tt2) Then
        tt3 = Split(tt2, ",")(0)          '获取联赛ID
        '判断联赛ID是否在要取的联赛ID中
        If tt3 <> "" Then
            itemId = tt(i).Cells(1).innerText    '联赛
            If leagueDict.exists(itemId) Then
                col = col + 1
                For j = 0 To tt(i).Cells.Length - 1
                    If j = colCnt Then
                        itemId = tt(i).Cells(j).ChildNodes(0).nameProp
                        data1(col, j) = Left(itemId, Len(itemId) - 4)
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

Set wkSheet = ActiveWorkbook.Worksheets(sheetName)
wkSheet.Cells.ClearContents
For i = 0 To col
    For j = 0 To colCnt + 7
        wkSheet.Cells(i + 1, j + 1) = data1(i, j)
    Next
Next

Set wkSheet = Nothing

End Sub




'防盗链数据抓取示例
Sub 中国竞彩网数据载入()
     Dim CookieHeaders, Cookie, winhttp As Object, i&
     Dim Cookie1
     Dim tt
     Dim tt1
     Dim tt2 As Object
     Dim tt3, tt4, tt5, tt6, tt7, tt8
     Dim j
     Dim isRun As Boolean
     Dim leagueData()
     Dim league
     Dim data()
     
     '取对应关系数据
     Call loadLeagueData(leagueData)
     
     '取网站数据
     Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")

     With winhttp
         .Option(6) = 0
         .Open "GET", "http://info.sporttery.cn/football/hhad_list.php", False              '执行这句，得到的网页数据是英文
        .setRequestHeader "Connection", "Keep-Alive"
         .send
         Cookie = "" 'Split(.getResponseHeader("Set-Cookie"), ";")(0)     '获取Cookie            '第一次获取Cookie，决定了语言

         
         Cookie = "Hm_lvt_860f3361e3ed9c994816101d37900758=1408275075,1408362165;" + Cookie + ";Hm_lpvt_860f3361e3ed9c994816101d37900758=1408364038"
         .Option(6) = 1
         .Option(2) = 936   '65001      ' 936或950或65001           'GB2312/BIG5/UTF-8
         .Open "GET", "http://i.sporttery.cn/odds_calculator/get_odds?i_format=json&i_callback=getData&poolcode[]=hhad&poolcode[]=had", False  '&_1408364398357", False          '第二次取得数据"
         .setRequestHeader "Referer", "http://info.sporttery.cn/football/hhad_list.php"
         .setRequestHeader "Cookie", Cookie
         .setRequestHeader "Connection", "Keep-Alive"
         .send
         '将二进制转换UTF-8
         tt = BytesToBstr(.responseBody, "UTF-8")
         '将UTF-8转换为汉字
         tt1 = UTF8toChineseCharacters(tt)
         tt = Mid(tt1, 9, Len(tt1) - 10)

        'With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")       '调试用，数据放入剪贴板
        '    .SetText tt 'Text
        '     .PutInClipboard
        'End With

         '将JSON数据转换为数组
         Call getItemfromJson(tt, tt2)
         
         Set tt3 = CallByName(tt2, "data", VbGet)
         
     End With
     tt4 = Split(tt, ":{""id"":""")
     ReDim data(UBound(tt4), 18)   'id,日期,编号，赛事，主队，客队，主胜，平，主负,主胜即时值，平即时值，主负即时值
     'add by ljqu    增加让球的定义：让球个数，主胜，平，主负，即时值主胜，平，主负
     For j = 1 To UBound(tt4)
        tt5 = Split(tt4(j), """,")
        tt6 = "_" + tt5(0)
        data(j, 0) = tt5(0)
        Set tt7 = CallByName(tt3, tt6, VbGet)
        data(j, 1) = CDate(CallByName(tt7, "date", VbGet))
        data(j, 2) = CallByName(tt7, "num", VbGet)
        league = CallByName(tt7, "l_cn_abbr", VbGet)   '赛事
        data(j, 3) = UniformLeague(leagueData, league, 4)
        data(j, 4) = CallByName(tt7, "h_cn", VbGet)
        data(j, 5) = CallByName(tt7, "a_cn", VbGet)
        On Error Resume Next
        Set tt8 = CallByName(tt7, "had", VbGet)
        If Not (tt8 Is Nothing) Then
            data(j, 6) = CDbl(CallByName(tt8, "h", VbGet))
            data(j, 7) = CDbl(CallByName(tt8, "d", VbGet))
            data(j, 8) = CDbl(CallByName(tt8, "a", VbGet))
            data(j, 9) = CDbl(CallByName(tt8, "h", VbGet))
            data(j, 10) = CDbl(CallByName(tt8, "d", VbGet))
            data(j, 11) = CDbl(CallByName(tt8, "a", VbGet))
        End If
        Set tt8 = CallByName(tt7, "hhad", VbGet)
        If Not (tt8 Is Nothing) Then
            data(j, 12) = CDbl(CallByName(tt8, "fixedodds", VbGet))
            data(j, 13) = CDbl(CallByName(tt8, "h", VbGet))
            data(j, 14) = CDbl(CallByName(tt8, "d", VbGet))
            data(j, 15) = CDbl(CallByName(tt8, "a", VbGet))
            data(j, 16) = CDbl(CallByName(tt8, "h", VbGet))
            data(j, 17) = CDbl(CallByName(tt8, "d", VbGet))
            data(j, 18) = CDbl(CallByName(tt8, "a", VbGet))
        End If
        
        
        Set tt7 = Nothing
        Set tt8 = Nothing
     Next
     
    
    Dim x1 As Worksheet
    Set x1 = ActiveWorkbook.Sheets("中国竞彩网")
    'cnt = x1.UsedRange.Rows(x1.UsedRange.Rows.Count).Row
    '清除原有内容
    x1.Cells.ClearContents
    
    x1.Select
    Selection.ClearContents
    x1.Cells(1, 1) = "ID"
    x1.Cells(1, 2) = "日期"
    x1.Cells(1, 3) = "编号"
    x1.Cells(1, 4) = "赛事"
    x1.Cells(1, 5) = "主队"
    x1.Cells(1, 6) = "客队"
    x1.Cells(1, 7) = "主胜"
    x1.Cells(1, 8) = "平"
    x1.Cells(1, 9) = "主负"
    x1.Cells(1, 10) = "即时值-主胜"
    x1.Cells(1, 11) = "平"
    x1.Cells(1, 12) = "主负"
    x1.Cells(1, 13) = "让球个数"
    x1.Cells(1, 14) = "主胜"
    x1.Cells(1, 15) = "平"
    x1.Cells(1, 16) = "主负"
    x1.Cells(1, 17) = "让即时值-主胜"
    x1.Cells(1, 18) = "平"
    x1.Cells(1, 19) = "主负"
    For i = 1 To UBound(data)
        '记录数据
        For j = 0 To UBound(data, 2)
            x1.Cells(i + 1, j + 1) = data(i, j)
        Next
    Next

     Set x1 = Nothing
     Set tt2 = Nothing
     Set tt3 = Nothing
     Set winhttp = Nothing
 End Sub



Function 取球队联赛积分(dataAvg, ids As String)
'------------------------------------------------------------------
'create 2015.3.28  ljqu
'
'dataAvg:返回主队和客队的联赛积分情况
'ids:球赛对应的id号
'------------------------------------------------------------------

Dim IE As Object
Dim doc As Object
Dim k As Integer
Dim i As Integer
Dim j As Integer
Dim data1(7, 10)
Dim rowcnt As Integer
Dim colCnt As Integer

Dim tt As Object
Dim tt1, tt2, tt3
Dim URL

URL = "http://zq.win007.com/analysis/" + ids + "cn.htm"

Set IE = UserForm1.WebBrowser1
With IE
  .Navigate URL '网址
  
  Do Until .ReadyState = 4
    DoEvents
  Loop
  Set doc = .document
End With
'Application.ScreenUpdating = False
ReDim dataAvg(103)   '0-43,主队积分信息，44-87：客队积分信息，
                   '88—91：Bet365欧转亚盘初盘：主队、让球、客队、总水位
                   '92—95：Bet365欧转亚盘即时：主队、让球、客队、总水位
                   '96—99：澳彩欧转亚盘初盘：主队、让球、客队、总水位
                   '100—103：澳彩欧转亚盘即时：主队、让球、客队、总水位
rowcnt = 7
colCnt = 10

    If doc.getElementById("porlet_5") Is Nothing Then
        取球队联赛积分 = False
        Exit Function
    End If
    Set tt = doc.getElementById("porlet_5").ChildNodes(0).ChildNodes(1)   '取联赛积分排名数据
    Set tt1 = tt.Cells(0).ChildNodes(0).ChildNodes(0)     '主队全场数据，保存：总、主、客、近6
    Set tt2 = tt.Cells(1).ChildNodes(0).ChildNodes(0)     '客队全场数据，保存：总、主、客、近6
    If tt1.ChildNodes.Length = 6 Then
       For i = 2 To tt1.ChildNodes.Length - 1
           For j = 0 To colCnt
            data1(i - 2, j) = tt1.ChildNodes(i).Cells(j).innerText
           Next
       Next
       For i = 2 To tt2.ChildNodes.Length - 1
           For j = 0 To colCnt
            data1(4 + i - 2, j) = tt2.ChildNodes(i).Cells(j).innerText
           Next
       Next

     
     '移动联赛积分数据
     For i = 0 To rowcnt
        For j = 0 To colCnt
            dataAvg(j + i * (colCnt + 1)) = data1(i, j)
        Next
     Next
     
     '获取即时赔率比较
     If doc.getElementById("porlet_1").ChildNodes.Length > 0 Then
     
        Set tt = doc.getElementById("porlet_1").ChildNodes(1).ChildNodes(0)   '取联赛积分排名数据
        
        For i = 0 To tt.Rows.Length - 1    '行
        
               If tt.Rows.Item(i) <> "" Then
                   tt3 = tt.Rows.Item(i).Cells(0).innerText
                   If InStr(tt3, "Bet365") > 0 Then
                       dataAvg(88) = tt.Rows.Item(i).Cells(5).innerText
                       dataAvg(89) = tt.Rows.Item(i).Cells(6).innerText
                       dataAvg(90) = tt.Rows.Item(i).Cells(7).innerText
                       dataAvg(91) = tt.Rows.Item(i).Cells(8).innerText
                       
                       dataAvg(92) = tt.Rows.Item(i + 1).Cells(4).innerText
                       dataAvg(93) = tt.Rows.Item(i + 1).Cells(5).innerText
                       dataAvg(94) = tt.Rows.Item(i + 1).Cells(6).innerText
                       dataAvg(95) = tt.Rows.Item(i + 1).Cells(7).innerText
                       
                   ElseIf InStr(tt3, "澳彩") > 0 Then
                       dataAvg(96) = tt.Rows.Item(i).Cells(5).innerText
                       dataAvg(97) = tt.Rows.Item(i).Cells(6).innerText
                       dataAvg(98) = tt.Rows.Item(i).Cells(7).innerText
                       dataAvg(99) = tt.Rows.Item(i).Cells(8).innerText
                       
                       dataAvg(100) = tt.Rows.Item(i + 1).Cells(4).innerText
                       dataAvg(101) = tt.Rows.Item(i + 1).Cells(5).innerText
                       dataAvg(102) = tt.Rows.Item(i + 1).Cells(6).innerText
                       dataAvg(103) = tt.Rows.Item(i + 1).Cells(7).innerText
                       
                   End If
               End If
        Next
     End If
     取球队联赛积分 = True
   Else
      取球队联赛积分 = False
   End If

     Set doc = Nothing
     Set IE = Nothing
End Function


Sub 球探网赛事积分数据载入()
'------------------------------------------------------------
'dataBF:数据输出的数组
'dataW:要查找的数据根据data(,0)的ID号去链接新的网址数据
'------------------------------------------------------------
Dim rowNo
Dim col
Dim i, j
Dim vsId As String
Dim data()
Dim bfData()
Dim srcdata()
Dim wkSheet As Worksheet

Call LoadDataToArray(srcdata, "球探网(W)")

rowNo = UBound(srcdata, 1)
ReDim data(rowNo, 104)     'id号，主队积分44个数据，客队积分44个数据
    '1-11:主队总赛事信息
    '12-22:主队主赛事信息
    '23-33:主队客赛事信息
    '34-44:主队近6场赛事信息
    '45-88:客队相关信息，也按照主队的规则进行排列
    '89—92：Bet365欧转亚盘初盘：主队、让球、客队、总水位
    '93—96：Bet365欧转亚盘即时：主队、让球、客队、总水位
    '97—100：澳彩欧转亚盘初盘：主队、让球、客队、总水位
    '101—104：澳彩欧转亚盘即时：主队、让球、客队、总水位
For i = 1 To rowNo
    vsId = srcdata(i, 0)
    'Sleep 80
    
    '根据vsId获取相应的数据组，四个指标数据
    'Call 取欧赔指数(bfData, CStr(vsId))
    
    If 取球队联赛积分(bfData, CStr(vsId)) Then     '如果有数据
    '拼装成新形式的数据格式
        data(i, 0) = vsId
        For j = 0 To 103
            data(i, j + 1) = bfData(j)
        Next
        
    End If
Next

Set wkSheet = ActiveWorkbook.Sheets("球探网积分")
wkSheet.Cells.ClearContents

wkSheet.Cells(1, 1) = "序号"
For i = 0 To 7
    If i < 4 Then
        wkSheet.Cells(1, i * 11 + 2) = "主队全场"
    Else
        wkSheet.Cells(1, i * 11 + 2) = "客队全场"
    End If
    wkSheet.Cells(1, i * 11 + 3) = "赛"
    wkSheet.Cells(1, i * 11 + 4) = "胜"
    wkSheet.Cells(1, i * 11 + 5) = "平"
    wkSheet.Cells(1, i * 11 + 6) = "负"
    wkSheet.Cells(1, i * 11 + 7) = "得"
    wkSheet.Cells(1, i * 11 + 8) = "失"
    wkSheet.Cells(1, i * 11 + 9) = "净"
    wkSheet.Cells(1, i * 11 + 10) = "得分"
    wkSheet.Cells(1, i * 11 + 11) = "排名"
    wkSheet.Cells(1, i * 11 + 12) = "胜率"
Next

wkSheet.Cells(1, 90) = "B初-主队"
wkSheet.Cells(1, 91) = "B初-让球"
wkSheet.Cells(1, 92) = "B初-客队"
wkSheet.Cells(1, 93) = "B初-总水位"
wkSheet.Cells(1, 94) = "B即-主队"
wkSheet.Cells(1, 95) = "B即-让球"
wkSheet.Cells(1, 96) = "B即-客队"
wkSheet.Cells(1, 97) = "B即-总水位"


wkSheet.Cells(1, 98) = "M初-主队"
wkSheet.Cells(1, 99) = "M初-让球"
wkSheet.Cells(1, 100) = "M初-客队"
wkSheet.Cells(1, 101) = "M初-总水位"
wkSheet.Cells(1, 102) = "M即-主队"
wkSheet.Cells(1, 103) = "M即-让球"
wkSheet.Cells(1, 104) = "M即-客队"
wkSheet.Cells(1, 105) = "M即-总水位"


For i = 1 To rowNo
    For j = 0 To 104
        wkSheet.Cells(i + 1, j + 1) = data(i, j)
    Next
Next

End Sub



'防盗链数据抓取示例
Sub 竞彩网比分数据载入()
     Dim CookieHeaders, Cookie, winhttp As Object, i&
     Dim Cookie1
     Dim tt
     Dim tt1
     Dim tt2 As Object
     Dim tt3, tt4, tt5, tt6, tt7, tt8
     Dim j
     Dim isRun As Boolean
     Dim leagueData()
     Dim league
     Dim data()
     
     '取对应关系数据
     Call loadLeagueData(leagueData)
     
     '取网站数据
     Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")

     With winhttp
         .Option(6) = 0
         .Open "GET", "http://info.sporttery.cn/football/cal_crs.htm", False              '执行这句，得到的网页数据是英文
        .setRequestHeader "Connection", "Keep-Alive"
         .send
         'Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)     '获取Cookie            '第一次获取Cookie，决定了语言

         
         Cookie = "Hm_lvt_860f3361e3ed9c994816101d37900758=1408275075,1408362165;" + Cookie + ";Hm_lpvt_860f3361e3ed9c994816101d37900758=1408364038"
         .Option(6) = 1
         .Option(2) = 936   '65001      ' 936或950或65001           'GB2312/BIG5/UTF-8
         .Open "GET", "http://i.sporttery.cn/odds_calculator/get_odds?i_format=json&i_callback=getData&poolcode[]=crs", False  '&_1408364398357", False          '第二次取得数据"
         .setRequestHeader "Referer", "http://info.sporttery.cn/football/cal_crs.htm"
         .setRequestHeader "Cookie", Cookie
         .setRequestHeader "Connection", "Keep-Alive"
         .send
         '将二进制转换UTF-8
         tt = BytesToBstr(.responseBody, "UTF-8")
         '将UTF-8转换为汉字
         tt1 = UTF8toChineseCharacters(tt)
         tt = Mid(tt1, 9, Len(tt1) - 10)

        'With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")       '调试用，数据放入剪贴板
        '    .SetText tt 'Text
        '     .PutInClipboard
         'End With

         '将JSON数据转换为数组
         Call getItemfromJson(tt, tt2)
         
         Set tt3 = CallByName(tt2, "data", VbGet)
         
     End With
     tt4 = Split(tt, ":{""id"":""")
     ReDim data(UBound(tt4), 36)   'id,日期,编号，赛事，主队，客队，
                                   '
     For j = 1 To UBound(tt4)
        tt5 = Split(tt4(j), """,")
        tt6 = "_" + tt5(0)
        data(j, 0) = tt5(0)
        Set tt7 = CallByName(tt3, tt6, VbGet)
        data(j, 1) = CDate(CallByName(tt7, "date", VbGet))
        data(j, 2) = CallByName(tt7, "num", VbGet)
        league = CallByName(tt7, "l_cn_abbr", VbGet)   '赛事
        data(j, 3) = UniformLeague(leagueData, league, 4)
        data(j, 4) = CallByName(tt7, "h_cn", VbGet)
        data(j, 5) = CallByName(tt7, "a_cn", VbGet)
        On Error Resume Next
        Set tt8 = CallByName(tt7, "crs", VbGet)
        If Not (tt8 Is Nothing) Then
            data(j, 6) = CDbl(CallByName(tt8, "0100", VbGet))
            data(j, 7) = CDbl(CallByName(tt8, "0200", VbGet))
            data(j, 8) = CDbl(CallByName(tt8, "0201", VbGet))
            data(j, 9) = CDbl(CallByName(tt8, "0300", VbGet))
            data(j, 10) = CDbl(CallByName(tt8, "0301", VbGet))
            data(j, 11) = CDbl(CallByName(tt8, "0302", VbGet))
            data(j, 12) = CDbl(CallByName(tt8, "0400", VbGet))
            data(j, 13) = CDbl(CallByName(tt8, "0401", VbGet))
            data(j, 14) = CDbl(CallByName(tt8, "0402", VbGet))
            data(j, 15) = CDbl(CallByName(tt8, "0500", VbGet))
            data(j, 16) = CDbl(CallByName(tt8, "0501", VbGet))
            data(j, 17) = CDbl(CallByName(tt8, "0502", VbGet))
            data(j, 18) = CDbl(CallByName(tt8, "-1-h", VbGet))
            
            data(j, 19) = CDbl(CallByName(tt8, "0000", VbGet))
            data(j, 20) = CDbl(CallByName(tt8, "0101", VbGet))
            data(j, 21) = CDbl(CallByName(tt8, "0202", VbGet))
            data(j, 22) = CDbl(CallByName(tt8, "0303", VbGet))
            data(j, 23) = CDbl(CallByName(tt8, "-1-d", VbGet))
            
            
            data(j, 24) = CDbl(CallByName(tt8, "0001", VbGet))
            data(j, 25) = CDbl(CallByName(tt8, "0002", VbGet))
            data(j, 26) = CDbl(CallByName(tt8, "0102", VbGet))
            data(j, 27) = CDbl(CallByName(tt8, "0003", VbGet))
            data(j, 28) = CDbl(CallByName(tt8, "0103", VbGet))
            data(j, 29) = CDbl(CallByName(tt8, "0203", VbGet))
            data(j, 30) = CDbl(CallByName(tt8, "0004", VbGet))
            data(j, 31) = CDbl(CallByName(tt8, "0104", VbGet))
            data(j, 32) = CDbl(CallByName(tt8, "0204", VbGet))
            data(j, 33) = CDbl(CallByName(tt8, "0005", VbGet))
            data(j, 34) = CDbl(CallByName(tt8, "0105", VbGet))
            data(j, 35) = CDbl(CallByName(tt8, "0205", VbGet))
            data(j, 36) = CDbl(CallByName(tt8, "-1-a", VbGet))
            
        End If
        
        
        
        Set tt7 = Nothing
        Set tt8 = Nothing
     Next
     
    
    Dim x1 As Worksheet
    Set x1 = ActiveWorkbook.Sheets("竞彩网比分")
    'cnt = x1.UsedRange.Rows(x1.UsedRange.Rows.Count).Row
    '清除原有内容
    x1.Cells.ClearContents
    
    x1.Select
    Selection.ClearContents
    x1.Cells(1, 1) = "ID"
    x1.Cells(1, 2) = "日期"
    x1.Cells(1, 3) = "编号"
    x1.Cells(1, 4) = "赛事"
    x1.Cells(1, 5) = "主队"
    x1.Cells(1, 6) = "客队"
    
    x1.Cells(1, 7) = "(胜)1:0"
    x1.Cells(1, 8) = "2:0"
    x1.Cells(1, 9) = "2:1"
    x1.Cells(1, 10) = "3:0"
    x1.Cells(1, 11) = "3:1"
    x1.Cells(1, 12) = "3:2"
    x1.Cells(1, 13) = "4:0"
    x1.Cells(1, 14) = "4:1"
    x1.Cells(1, 15) = "4:2"
    x1.Cells(1, 16) = "5:0"
    x1.Cells(1, 17) = "5:1"
    x1.Cells(1, 18) = "5:2"
    x1.Cells(1, 19) = "胜其他"
    
    
    x1.Cells(1, 20) = "(平)0:0"
    x1.Cells(1, 21) = "1:1"
    x1.Cells(1, 22) = "2:2"
    x1.Cells(1, 23) = "3:3"
    x1.Cells(1, 24) = "平其他"
    
    
    x1.Cells(1, 25) = "(负)0:1"
    x1.Cells(1, 26) = "0:2"
    x1.Cells(1, 27) = "1:2"
    x1.Cells(1, 28) = "0:3"
    x1.Cells(1, 29) = "1:3"
    x1.Cells(1, 30) = "2:3"
    x1.Cells(1, 31) = "0:4"
    x1.Cells(1, 32) = "1:4"
    x1.Cells(1, 33) = "2:4"
    x1.Cells(1, 34) = "0:5"
    x1.Cells(1, 35) = "1:5"
    x1.Cells(1, 36) = "2:5"
    x1.Cells(1, 37) = "负其他"
    
    
    For i = 1 To UBound(data)
        '记录数据
        For j = 0 To UBound(data, 2)
            x1.Cells(i + 1, j + 1) = data(i, j)
        Next
    Next

     Set x1 = Nothing
     Set tt2 = Nothing
     Set tt3 = Nothing
     Set winhttp = Nothing
 End Sub
 
 
 
 '防盗链数据抓取示例
Sub 竞彩网总进球数据载入()
     Dim CookieHeaders, Cookie, winhttp As Object, i&
     Dim Cookie1
     Dim tt
     Dim tt1
     Dim tt2 As Object
     Dim tt3, tt4, tt5, tt6, tt7, tt8
     Dim j
     Dim isRun As Boolean
     Dim leagueData()
     Dim league
     Dim data()
     
     '取对应关系数据
     Call loadLeagueData(leagueData)
     
     '取网站数据
     Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")

     With winhttp
         .Option(6) = 0
         .Open "GET", "http://info.sporttery.cn/football/cal_ttg.htm", False              '执行这句，得到的网页数据是英文
        .setRequestHeader "Connection", "Keep-Alive"
         .send
         'Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)     '获取Cookie            '第一次获取Cookie，决定了语言

         
         Cookie = "Hm_lvt_860f3361e3ed9c994816101d37900758=1408275075,1408362165;" + Cookie + ";Hm_lpvt_860f3361e3ed9c994816101d37900758=1408364038"
         .Option(6) = 1
         .Option(2) = 936   '65001      ' 936或950或65001           'GB2312/BIG5/UTF-8
         .Open "GET", "http://i.sporttery.cn/odds_calculator/get_odds?i_format=json&i_callback=getData&poolcode[]=ttg", False  '&_1408364398357", False          '第二次取得数据"
         .setRequestHeader "Referer", "http://info.sporttery.cn/football/cal_ttg.htm"
         .setRequestHeader "Cookie", Cookie
         .setRequestHeader "Connection", "Keep-Alive"
         .send
         '将二进制转换UTF-8
         tt = BytesToBstr(.responseBody, "UTF-8")
         '将UTF-8转换为汉字
         tt1 = UTF8toChineseCharacters(tt)
         tt = Mid(tt1, 9, Len(tt1) - 10)

        'With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")       '调试用，数据放入剪贴板
        '    .SetText tt 'Text
        '     .PutInClipboard
         'End With

         '将JSON数据转换为数组
         Call getItemfromJson(tt, tt2)
         
         Set tt3 = CallByName(tt2, "data", VbGet)
         
     End With
     tt4 = Split(tt, ":{""id"":""")
     ReDim data(UBound(tt4), 13)   'id,日期,编号，赛事，主队，客队，0球，1球，2球，3球，4球，5球，6球，7+球

     For j = 1 To UBound(tt4)
        tt5 = Split(tt4(j), """,")
        tt6 = "_" + tt5(0)
        data(j, 0) = tt5(0)
        Set tt7 = CallByName(tt3, tt6, VbGet)
        data(j, 1) = CDate(CallByName(tt7, "date", VbGet))
        data(j, 2) = CallByName(tt7, "num", VbGet)
        league = CallByName(tt7, "l_cn_abbr", VbGet)   '赛事
        data(j, 3) = UniformLeague(leagueData, league, 4)
        data(j, 4) = CallByName(tt7, "h_cn", VbGet)
        data(j, 5) = CallByName(tt7, "a_cn", VbGet)
        On Error Resume Next
        Set tt8 = CallByName(tt7, "ttg", VbGet)
        If Not (tt8 Is Nothing) Then
            data(j, 6) = CDbl(CallByName(tt8, "s0", VbGet))
            data(j, 7) = CDbl(CallByName(tt8, "s1", VbGet))
            data(j, 8) = CDbl(CallByName(tt8, "s2", VbGet))
            data(j, 9) = CDbl(CallByName(tt8, "s3", VbGet))
            data(j, 10) = CDbl(CallByName(tt8, "s4", VbGet))
            data(j, 11) = CDbl(CallByName(tt8, "s5", VbGet))
            data(j, 12) = CDbl(CallByName(tt8, "s6", VbGet))
            data(j, 13) = CDbl(CallByName(tt8, "s7", VbGet))
        End If

        
        
        Set tt7 = Nothing
        Set tt8 = Nothing
     Next
     
    
    Dim x1 As Worksheet
    Set x1 = ActiveWorkbook.Sheets("竞彩网总进球")
    'cnt = x1.UsedRange.Rows(x1.UsedRange.Rows.Count).Row
    '清除原有内容
    x1.Cells.ClearContents
    
    x1.Select
    Selection.ClearContents
    x1.Cells(1, 1) = "ID"
    x1.Cells(1, 2) = "日期"
    x1.Cells(1, 3) = "编号"
    x1.Cells(1, 4) = "赛事"
    x1.Cells(1, 5) = "主队"
    x1.Cells(1, 6) = "客队"
    x1.Cells(1, 7) = "0球"
    x1.Cells(1, 8) = "1球"
    x1.Cells(1, 9) = "2球"
    x1.Cells(1, 10) = "3球"
    x1.Cells(1, 11) = "4球"
    x1.Cells(1, 12) = "5球"
    x1.Cells(1, 13) = "6球"
    x1.Cells(1, 14) = "7+球"

    For i = 1 To UBound(data)
        '记录数据
        For j = 0 To UBound(data, 2)
            x1.Cells(i + 1, j + 1) = data(i, j)
        Next
    Next

     Set x1 = Nothing
     Set tt2 = Nothing
     Set tt3 = Nothing
     Set winhttp = Nothing
 End Sub
 
 
 '防盗链数据抓取示例
Sub 竞彩网半全场胜平负数据载入()
     Dim CookieHeaders, Cookie, winhttp As Object, i&
     Dim Cookie1
     Dim tt
     Dim tt1
     Dim tt2 As Object
     Dim tt3, tt4, tt5, tt6, tt7, tt8
     Dim j
     Dim isRun As Boolean
     Dim leagueData()
     Dim league
     Dim data()
     
     
     
     '取对应关系数据
     Call loadLeagueData(leagueData)
     
     '取网站数据
     Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")

     With winhttp
         .Option(6) = 0
         .Open "GET", "http://info.sporttery.cn/football/cal_hafu.htm", False              '执行这句，得到的网页数据是英文
        .setRequestHeader "Connection", "Keep-Alive"
         .send
         'Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)     '获取Cookie            '第一次获取Cookie，决定了语言

         
         Cookie = "Hm_lvt_860f3361e3ed9c994816101d37900758=1408275075,1408362165;" + Cookie + ";Hm_lpvt_860f3361e3ed9c994816101d37900758=1408364038"
         .Option(6) = 1
         .Option(2) = 936   '65001      ' 936或950或65001           'GB2312/BIG5/UTF-8
         .Open "GET", "http://i.sporttery.cn/odds_calculator/get_odds?i_format=json&i_callback=getData&poolcode[]=hafu", False  '&_1408364398357", False          '第二次取得数据"
         .setRequestHeader "Referer", "http://info.sporttery.cn/football/cal_hafu.htm"
         .setRequestHeader "Cookie", Cookie
         .setRequestHeader "Connection", "Keep-Alive"
         .send
         '将二进制转换UTF-8
         tt = BytesToBstr(.responseBody, "UTF-8")
         '将UTF-8转换为汉字
         tt1 = UTF8toChineseCharacters(tt)
         tt = Mid(tt1, 9, Len(tt1) - 10)

        'With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")       '调试用，数据放入剪贴板
        '    .SetText tt 'Text
        '     .PutInClipboard
         'End With

         '将JSON数据转换为数组
         Call getItemfromJson(tt, tt2)
         
         Set tt3 = CallByName(tt2, "data", VbGet)
         
     End With
     tt4 = Split(tt, ":{""id"":""")
     ReDim data(UBound(tt4), 14)   'id,日期,编号，赛事，主队，客队，胜胜，胜平，胜平，平胜，平平，平负，负胜，负平，负负
     For j = 1 To UBound(tt4)
        tt5 = Split(tt4(j), """,")
        tt6 = "_" + tt5(0)
        data(j, 0) = tt5(0)
        Set tt7 = CallByName(tt3, tt6, VbGet)
        data(j, 1) = CDate(CallByName(tt7, "date", VbGet))
        data(j, 2) = CallByName(tt7, "num", VbGet)
        league = CallByName(tt7, "l_cn_abbr", VbGet)   '赛事
        data(j, 3) = UniformLeague(leagueData, league, 4)
        data(j, 4) = CallByName(tt7, "h_cn", VbGet)
        data(j, 5) = CallByName(tt7, "a_cn", VbGet)
        On Error Resume Next
        Set tt8 = CallByName(tt7, "hafu", VbGet)
        If Not (tt8 Is Nothing) Then
            data(j, 6) = CDbl(CallByName(tt8, "hh", VbGet))
            data(j, 7) = CDbl(CallByName(tt8, "hd", VbGet))
            data(j, 8) = CDbl(CallByName(tt8, "ha", VbGet))
            data(j, 9) = CDbl(CallByName(tt8, "dh", VbGet))
            data(j, 10) = CDbl(CallByName(tt8, "dd", VbGet))
            data(j, 11) = CDbl(CallByName(tt8, "da", VbGet))
            data(j, 12) = CDbl(CallByName(tt8, "ah", VbGet))
            data(j, 13) = CDbl(CallByName(tt8, "ad", VbGet))
            data(j, 14) = CDbl(CallByName(tt8, "aa", VbGet))
            
        End If
        
        
        Set tt7 = Nothing
        Set tt8 = Nothing
     Next
     
    
    Dim x1 As Worksheet
    Set x1 = ActiveWorkbook.Sheets("竞彩网半全场")
    'cnt = x1.UsedRange.Rows(x1.UsedRange.Rows.Count).Row
    '清除原有内容
    x1.Cells.ClearContents
    
    x1.Select
    Selection.ClearContents
    x1.Cells(1, 1) = "ID"
    x1.Cells(1, 2) = "日期"
    x1.Cells(1, 3) = "编号"
    x1.Cells(1, 4) = "赛事"
    x1.Cells(1, 5) = "主队"
    x1.Cells(1, 6) = "客队"
    x1.Cells(1, 7) = "胜胜"
    x1.Cells(1, 8) = "胜平"
    x1.Cells(1, 9) = "胜负"
    x1.Cells(1, 10) = "平胜"
    x1.Cells(1, 11) = "平平"
    x1.Cells(1, 12) = "平负"
    x1.Cells(1, 13) = "负胜"
    x1.Cells(1, 14) = "负平"
    x1.Cells(1, 15) = "负负"

    For i = 1 To UBound(data)
        '记录数据
        For j = 0 To UBound(data, 2)
            x1.Cells(i + 1, j + 1) = data(i, j)
        Next
    Next

     Set x1 = Nothing
     Set tt2 = Nothing
     Set tt3 = Nothing
     Set winhttp = Nothing
 End Sub
 
 
 
 '防盗链数据抓取示例
Sub 竞彩网投资比例数据载入()
     Dim CookieHeaders, Cookie, winhttp As Object, i&
     Dim Cookie1
     Dim tt
     Dim tt1
     Dim tt2 As Object
     Dim tt3, tt4, tt5, tt6, tt7, tt8
     Dim j
     Dim isRun As Boolean
     Dim leagueData()
     Dim league
     Dim data()
     
     Dim tempDict As Object    '临时字典
     
     Dim x1 As Worksheet
     Dim x2 As Worksheet
     Dim Loc As Long              '位置
    
     
     '取对应关系数据
     Call loadLeagueData(leagueData)
     
     '取网站数据
     Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")

     With winhttp
         .Option(6) = 0
         .Open "GET", "http://info.sporttery.cn/football/hhad_list.php", False              '执行这句，得到的网页数据是英文
        .setRequestHeader "Connection", "Keep-Alive"
         .send
         'Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)     '获取Cookie            '第一次获取Cookie，决定了语言

         
         Cookie = "Hm_lvt_860f3361e3ed9c994816101d37900758=1468587314,1470742060;" + Cookie + ";Hm_lpvt_860f3361e3ed9c994816101d37900758=1470973763"
         .Option(6) = 1
         .Option(2) = 936   '65001      ' 936或950或65001           'GB2312/BIG5/UTF-8
         .Open "GET", "http://i.sporttery.cn/odds_calculator/get_proportion?i_format=json&i_callback=getReferData1&poolcode[]=hhad&poolcode[]=had", False  '&_1408364398357", False          '第二次取得数据"
         .setRequestHeader "Referer", "http://info.sporttery.cn/football/hhad_list.php"
         .setRequestHeader "Cookie", Cookie
         .setRequestHeader "Connection", "Keep-Alive"
         .send
         '将二进制转换UTF-8
         tt = BytesToBstr(.responseBody, "UTF-8")
         '将UTF-8转换为汉字
         tt1 = UTF8toChineseCharacters(tt)
         tt = Mid(tt1, 15, Len(tt1) - 16)

        'With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")       '调试用，数据放入剪贴板
        '    .SetText tt 'Text
        '     .PutInClipboard
         'End With

         '将JSON数据转换为数组
         Call getItemfromJson(tt, tt2)
         
         Set tt3 = CallByName(tt2, "data", VbGet)
         
     End With
     tt4 = Split(tt, """_")
     ReDim data(UBound(tt4), 18)   'id,日期,编号，赛事，主队，客队，主胜，平，主负,百分比主胜，平，主负，让-主胜，平，主负，让百分比-主胜，平，主负
     For j = 1 To UBound(tt4)
        tt5 = Split(tt4(j), """:{""")
        tt6 = "_" + tt5(0)
        data(j, 0) = CDbl(tt5(0))
        Set tt7 = CallByName(tt3, tt6, VbGet)
        
        'data(j, 2) = CallByName(tt7, "num", VbGet)
        'league = CallByName(tt7, "l_cn_abbr", VbGet)   '赛事
        'data(j, 3) = UniformLeague(leagueData, league, 4)
        'data(j, 4) = CallByName(tt7, "h_cn", VbGet)
        'data(j, 5) = CallByName(tt7, "a_cn", VbGet)
        On Error Resume Next
        Set tt8 = CallByName(tt7, "had", VbGet)
        If Not (tt8 Is Nothing) Then
            data(j, 1) = CDate(CallByName(tt8, "num", VbGet))
            data(j, 6) = CDbl(CallByName(tt8, "win", VbGet))
            data(j, 7) = CDbl(CallByName(tt8, "draw", VbGet))
            data(j, 8) = CDbl(CallByName(tt8, "lose", VbGet))
            data(j, 9) = CallByName(tt8, "pre_win", VbGet)
            data(j, 10) = CallByName(tt8, "pre_draw", VbGet)
            data(j, 11) = CallByName(tt8, "pre_lose", VbGet)
        End If
        Set tt8 = CallByName(tt7, "hhad", VbGet)
        If Not (tt8 Is Nothing) Then
            data(j, 12) = CDbl(CallByName(tt8, "fixedodds", VbGet))
            data(j, 13) = CDbl(CallByName(tt8, "win", VbGet))
            data(j, 14) = CDbl(CallByName(tt8, "draw", VbGet))
            data(j, 15) = CDbl(CallByName(tt8, "lose", VbGet))
            data(j, 16) = CallByName(tt8, "pre_win", VbGet)
            data(j, 17) = CallByName(tt8, "pre_draw", VbGet)
            data(j, 18) = CallByName(tt8, "pre_lose", VbGet)
        End If
        
        
        Set tt7 = Nothing
        Set tt8 = Nothing
     Next
     
    
    Set x2 = ActiveWorkbook.Sheets("中国竞彩网")
    x2.Cells(1, 20) = "让球主胜"
    x2.Cells(1, 21) = "平"
    x2.Cells(1, 22) = "主负"
    Call 初始化一般字典(tempDict, x2, 1, 0, 2, True)
    
    
    Set x1 = ActiveWorkbook.Sheets("竞彩网投资")
    'cnt = x1.UsedRange.Rows(x1.UsedRange.Rows.Count).Row
    '清除原有内容
    x1.Cells.ClearContents
    

    
    x1.Select
    Selection.ClearContents
    x1.Cells(1, 1) = "ID"
    x1.Cells(1, 2) = "日期"
    x1.Cells(1, 3) = "编号"
    x1.Cells(1, 4) = "赛事"
    x1.Cells(1, 5) = "主队"
    x1.Cells(1, 6) = "客队"
    x1.Cells(1, 7) = "主胜"
    x1.Cells(1, 8) = "平"
    x1.Cells(1, 9) = "主负"
    x1.Cells(1, 10) = "百分比-主胜"
    x1.Cells(1, 11) = "平"
    x1.Cells(1, 12) = "主负"
    x1.Cells(1, 13) = "让球个数"
    x1.Cells(1, 14) = "主胜"
    x1.Cells(1, 15) = "平"
    x1.Cells(1, 16) = "主负"
    x1.Cells(1, 17) = "让百分比-主胜"
    x1.Cells(1, 18) = "平"
    x1.Cells(1, 19) = "主负"

    For i = 1 To UBound(data)
        '记录数据
        For j = 0 To UBound(data, 2)
            x1.Cells(i + 1, j + 1) = data(i, j)
        Next
        
        '同时将数据更新到《中国竞彩网》sheet页
        If tempDict.exists(data(i, 0)) Then
            Loc = CLng(tempDict.Item(data(i, 0)))
            x2.Cells(Loc, 20) = data(i, 16)
            x2.Cells(Loc, 21) = data(i, 17)
            x2.Cells(Loc, 22) = data(i, 18)
        End If
    Next

     Set x1 = Nothing
     Set tt2 = Nothing
     Set tt3 = Nothing
     Set winhttp = Nothing
 End Sub


