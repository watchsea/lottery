Attribute VB_Name = "deal"
Option Explicit


Sub 数据初始(ByRef control As Office.IRibbonControl)
'------------------------------------------------------------------------------------
'数据初始化分为两步走：
'    1.从【综合数据】第一次加载数据
'    2.根据威廉希尔的数据，对未记载的数据进行新增
'    3.第二次加载【综合数据】页的数据
'    4.处理其他类别的数据
'------------------------------------------------------------------------------------

Dim wkWorkbook As Workbook

Dim wkSheet As Worksheet
Dim dataSheet As Worksheet

Dim dataDict As Object     '综合数据字典

Dim data()
Dim dataW()
Dim dataB()
Dim dataM()
Dim dataL()
Dim dataE()
Dim dataJ()
Dim dataL12()   '赔1、赔2
Dim dataJF()    '球探网联赛积分
Dim dataOZ()     '欧指数据   2020.03.15
Dim dataYZ()    '亚指数据   2020.03.15
Dim datatmp()   '临时用数据表 2020.03.15
Dim datasrc As String '数据来源表名 2020.03.15


Dim dataBF()   '澳客网必发
Dim dataSF()   '澳客网胜负
Dim dataKL()   '澳客网凯利
Dim dataPK()   '澳客网盘口预测

Dim dataConfig()   '配置数据信息

Dim dataJBf()     '竞彩网比分
Dim dataJZjq()    '竞彩网总进球
Dim dataJBqc()    '竞彩网半全场胜平负


Dim i As Long
Dim j As Long
Dim k As Long
Dim k1 As Long
Dim Loc As Long
Dim loc1 As Long
Dim wloc As Long      '威廉数据的记录个数
Dim tloc As Long      '总的数据循环次数


'定义赛事信息
Dim league        '联赛
Dim currdate      '赛事日期
Dim vsInfo        '对阵信息
Dim priTeam       '主队
Dim secTeam       '客队
Dim vsId          '球探网赛事ID

'参数

Dim calMeth As String    '数据处理方式，1：按最小记录进行处理，9：按最大记录处理
Dim dataWcol As Integer   '威廉希尔数据开始列号
Dim dataBcol As Integer     'Bet365数据开始列号
Dim dataMcol As Integer     '澳门彩票数据开始列号
Dim dataLcol As Integer     '立博(英国)数据开始列号
Dim dataEcol As Integer     '易胜博数据开始列号
Dim DataJcol As Integer     '竞彩网数据开始列号
Dim LOSE1col As Integer     '赔1 数据开始列号
Dim lose2Col As Integer     '赔2 数据开始列号
Dim BF1Col   As Integer     'BF1 数据开始列号
Dim BF2col  As Integer      'BF2 数据开始列号
Dim BF3col  As Integer      'BF3 数据开始列号
Dim OKBF1col As Integer     'OKBF1数据开始列号
Dim OKBF2col As Integer     'OKBF2数据开始列号
Dim OKnetCurrentCol As Integer   '澳客网当期期数存放列
Dim OK30col As Integer      'Ok30数据——澳客网胜负数据
Dim BFWcol As Integer       '威廉希尔——澳客网凯利方差
Dim BFMcol As Integer       'Bet365——澳客网凯利方差
Dim BFBcol As Integer       '澳门——澳客网凯利方差
Dim varCol As Integer       '方差——澳客网凯利方差

Dim varCmpCol As Long       '方差分析数据开始列号
Dim dataJRqCol As Long      '竞彩网让球数据开始列号
Dim dataJBfCol As Long      '竞彩网比分数据开始列号
Dim dataJZjqCol As Long     '竞彩网总让球开始列号
Dim dataJBqcCol As Long     '竞彩网半全场胜平负数据开始列号


Dim dataBgCol As Long    '数据在SHEET对应的数据起始列
Dim initialBgCol As Integer   '初始数据在数组中的起始列
Dim realBgCol As Integer   '即时数据在数组中的起始列
Dim funName As String      '动态调用的函数


Dim isCollectOknet As Boolean   '是否采集澳客网数据
Dim isFind As Boolean           '是否找到相应的记录
Dim dealRecCount As Long        '综合数据加载到内存的记录条数

Dim cnt

Dim insertRecno As Long         '新记录插入的位置信息


Call 初始化字典(Dict, "Param")
Call 初始化字典(leagueDict, "01赛事")
Call 配置数据载入(dataConfig, "Config")

calMeth = Dict.Item("CALMETH")


dealRecCount = CLng(Dict.Item("DEAL_RECCOUNT"))
dataBgCol = CLng(Dict.Item("DATABGCOL"))


isCollectOknet = CBool(Dict.Item("IS_COLLECT_OKNET"))



'第一次将综合数据加载入内存

Call 综合数据载入内存(data, "综合数据", dealRecCount, dataBgCol)
tloc = UBound(data)

'将ID号和对应的行号载入字典,8列：球赛Id，0列：对应的excel行号
If Not 载入综合数据字典(dataDict, data, 8, 0, "初始值") Then
    MsgBox ("加载综合数据指针字典时出错！")
    Exit Sub
End If

Set dataSheet = ActiveWorkbook.Sheets("综合数据")
Call 初始化一般字典(dataColDict, dataSheet, 4, 0, 1, False)


'取各数据项对应的列号

dataWcol = dataColDict.Item("DATAW")
dataBcol = dataColDict.Item("DATAB")
dataMcol = dataColDict.Item("DATAM")
dataLcol = dataColDict.Item("DATAL")
dataEcol = dataColDict.Item("DATAE")
DataJcol = dataColDict.Item("DATAJ")

LOSE1col = dataColDict.Item("LOSE1")
lose2Col = dataColDict.Item("LOSE2")
BF1Col = dataColDict.Item("BF1")
BF2col = dataColDict.Item("BF2")
BF3col = dataColDict.Item("BF3")

OKBF1col = dataColDict.Item("OKBF1")
OKBF2col = dataColDict.Item("OKBF2")

OK30col = dataColDict.Item("OK30")
BFWcol = dataColDict.Item("BFW")
BFMcol = dataColDict.Item("BFM")
BFBcol = dataColDict.Item("BFB")
varCol = dataColDict.Item("VAR")

varCmpCol = dataColDict.Item("VARCMP")
dataJRqCol = dataColDict.Item("DATAJRQ")
dataJBfCol = dataColDict.Item("DATAJBF")
dataJZjqCol = dataColDict.Item("DATAJZJQ")
dataJBqcCol = dataColDict.Item("DATAJBQC")



OKnetCurrentCol = dataColDict.Item("OKID")



'一
Call LoadDataToArray(dataW, "球探网(W)")
wloc = UBound(dataW)

'依据威廉希尔的数据去追踪bf3数据、赔1数据、赔2数据、以及bf3的数据
Call BF数据载入(dataL12, "球探网(BF)")

'--------------------------------------------
'  从2014.10.6开始，B、M、L、E的数据从BF中取
'--------------------------------------------
Call BF载入赔率(dataB, "球探网(BF)", "B")      'bet365(英国)
Call BF载入赔率(dataM, "球探网(BF)", "M")       '澳门
Call BF载入赔率(dataL, "球探网(BF)", "L")       '立博（英国）
Call BF载入赔率(dataE, "球探网(BF)", "E")       '易胜博
Call 球探网联赛积分载入(dataJF, "球探网积分")    '联赛积分
Call 欧指数据载入(dataOZ, "OZ")                 '欧指数据 2020.03.15
Call 亚指数据载入(dataYZ, "YZ")                 '亚指数据 2020.03.15

'---------------------------------
'       澳客网数据载入
'---------------------------------
Call 澳客网必发盈亏载入(dataBF, "澳客网(1)")
Call 澳客网凯利方差载入(dataKL, "澳客网(3)")
Call 澳客网胜负指数载入(dataSF, "澳客网(2)")
Call 澳客网盘口评测载入(dataPK, "澳客网(4)")


Call 加载中国竞彩网数据(dataJ, "中国竞彩网")
Call 加载竞彩网比分数据(dataJBf, "竞彩网比分")
Call 加载竞彩网总进球数据(dataJZjq, "竞彩网总进球")
Call 加载竞彩网半全场胜平负数据(dataJBqc, "竞彩网半全场")


' 根据威廉希尔的记录，确定要新增的记录

  loc1 = data(tloc, 0) + 1
  For i = 1 To wloc
      j = 1         '此处是否考虑每次都及时更新data中的数据
      insertRecno = 0
      isFind = False
      Do While j < tloc
          If data(j, 8) = dataW(i, 0) Then '球探网赛事ID
              Exit Do
          Else
              If j = 1 And (data(j, 1) > dataW(i, 2) Or data(j, 1) > dataW(i, 2) And data(j, 2) > dataW(i, 3)) Then  '插入到第一个记录
              '定位插入的位置
                  Loc = data(j, 0)
                  insertRecno = j
                  isFind = True
              ElseIf ((data(j - 1, 1) <= dataW(i, 2) And data(j, 1) > dataW(i, 2)) Or (data(j, 1) = dataW(i, 2) And data(j - 1, 2) <= dataW(i, 3) And data(j, 2) > dataW(i, 3))) And Not isFind Then  '中间数据的插入
              '定位插入的位置
                  Loc = data(j, 0)
                  insertRecno = j
                  isFind = True
              End If
          End If
          
          '指针移动
          If data(j, 6) = "初始值" Then
              j = j + 3
          Else
              j = j + 1
          End If
      Loop
      
      If insertRecno = 0 Then    '如果是当前最新的赛事
          Loc = loc1
      Else
          '在SHEET中插入三行
          dataSheet.Cells(Loc, 1).Resize(3, 1).EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
          '修改自insertRecno以后的数据指针
          For j = insertRecno To tloc
              data(j, 0) = data(j, 0) + 3
          Next
      End If
 
      If j >= tloc Then    '没有找到相应的记录，插入新记录
          dataSheet.Cells(Loc, 1) = dataW(i, 2) '日期
          dataSheet.Cells(Loc, 2) = dataW(i, 3) '时间
          dataSheet.Cells(Loc, 3) = dataW(i, 4) '主队
          dataSheet.Cells(Loc, 4) = dataW(i, 5) '客队
          dataSheet.Cells(Loc, 5) = dataW(i, 6) '对阵
          dataSheet.Cells(Loc, 6) = "初始值" '初始值
          dataSheet.Cells(Loc, 7) = dataW(i, 1) '联赛
          '赛季的算法
          dataSheet.Cells(Loc, 8) = 赛季计算(dataW(i, 2)) '赛季
          dataSheet.Cells(Loc, 9) = dataW(i, 0)    '球赛ID
          
          
          '记录比赛结果
          If dataSheet.Cells(Loc, 10) = "" And dataW(i, 15) <> "" Then
              dataSheet.Cells(Loc, 10) = dataW(i, 15)
              dataSheet.Cells(Loc, 11) = 比赛结果(dataW(i, 15))
              
              dataSheet.Cells(Loc + 1, 10) = dataSheet.Cells(Loc, 10)   '比分
              dataSheet.Cells(Loc + 1, 11) = dataSheet.Cells(Loc, 11)   '赛果
              dataSheet.Cells(Loc + 2, 10) = dataSheet.Cells(Loc, 10)   '比分
              dataSheet.Cells(Loc + 2, 11) = dataSheet.Cells(Loc, 11)   '赛果
              
          End If
          
          dataSheet.Cells(Loc + 1, 1) = dataSheet.Cells(Loc, 1) '日期
          dataSheet.Cells(Loc + 1, 2) = dataSheet.Cells(Loc, 2) '时间
          dataSheet.Cells(Loc + 1, 3) = dataSheet.Cells(Loc, 3) '主队
          dataSheet.Cells(Loc + 1, 4) = dataSheet.Cells(Loc, 4) '客队
          dataSheet.Cells(Loc + 1, 5) = dataSheet.Cells(Loc, 5) '对阵
          dataSheet.Cells(Loc + 1, 6) = "即时值1"
          dataSheet.Cells(Loc + 1, 7) = dataSheet.Cells(Loc, 7) '联赛
          dataSheet.Cells(Loc + 1, 8) = dataSheet.Cells(Loc, 8) '赛季
          dataSheet.Cells(Loc + 1, 9) = dataSheet.Cells(Loc, 9)  '球赛ID

          
          dataSheet.Cells(Loc + 2, 1) = dataSheet.Cells(Loc, 1) '日期
          dataSheet.Cells(Loc + 2, 2) = dataSheet.Cells(Loc, 2) '时间
          dataSheet.Cells(Loc + 2, 3) = dataSheet.Cells(Loc, 3) '主队
          dataSheet.Cells(Loc + 2, 4) = dataSheet.Cells(Loc, 4) '客队
          dataSheet.Cells(Loc + 2, 5) = dataSheet.Cells(Loc, 5) '对阵
          dataSheet.Cells(Loc + 2, 6) = "即时值2" '数据类型
          dataSheet.Cells(Loc + 2, 7) = dataSheet.Cells(Loc, 7) '联赛
          dataSheet.Cells(Loc + 2, 8) = dataSheet.Cells(Loc, 8) '赛季
          dataSheet.Cells(Loc + 2, 9) = dataSheet.Cells(Loc, 9)  '球赛ID

          Call 记录初始值(dataSheet, dataW, Loc, dataWcol, i, 7, 11, True, False, "A", 4)
          
         '*******************************************************************
         '               竞彩网相关数据处理
         '********************************************************************
        
         '由于竞彩网中世界杯预选赛（世预赛）在球探网中对应了多个，如”亚洲预选、南美预选、欧洲预选等"
         '因而当联赛名称中包含"预选"字样时，将不再比对联赛名称，add by ljqu 2016.11.15
         league = dataW(i, 1)
         
         '中国竞彩网
         For k = 1 To UBound(dataJ)
             If dataW(i, 2) = dataJ(k, 1) And (dataW(i, 4) = dataJ(k, 4) Or dataW(i, 5) = dataJ(k, 5)) And (dataW(i, 1) = dataJ(k, 3) Or InStr(league, "预选") > 0) Then '日期，主队，客队，联赛
                 Exit For
             End If
         Next
         
         If k <= UBound(dataJ) Then
             Call 记录初始值(dataSheet, dataJ, Loc, DataJcol, k, 6, 9, True, True, "D", 3)
         End If
         
         '中国竞彩网让球
         For k = 1 To UBound(dataJ)
             If dataW(i, 2) = dataJ(k, 1) And (dataW(i, 4) = dataJ(k, 4) Or dataW(i, 5) = dataJ(k, 5)) And (dataW(i, 1) = dataJ(k, 3) Or InStr(league, "预选") > 0) Then  '日期，主队，客队，联赛
                 Exit For
             End If
         Next
         
         If k <= UBound(dataJ) Then
             Call 记录初始值(dataSheet, dataJ, Loc, dataJRqCol, k, 19, 19, True, True, "D", 3)
         End If
         
         '中国竞彩网比分
         For k = 1 To UBound(dataJBf)
             If dataW(i, 2) = dataJBf(k, 1) And (dataW(i, 4) = dataJBf(k, 4) Or dataW(i, 5) = dataJBf(k, 5)) And (dataW(i, 1) = dataJBf(k, 3) Or InStr(league, "预选") > 0) Then  '日期，主队，客队，联赛
                 Exit For
             End If
         Next
         
         If k <= UBound(dataJBf) Then
             Call 记录初始值(dataSheet, dataJBf, Loc, dataJBfCol, k, 6, 6, True, False, "-", 31)
         End If
         
         '中国竞彩网总进球
         For k = 1 To UBound(dataJZjq)
             If dataW(i, 2) = dataJZjq(k, 1) And (dataW(i, 4) = dataJZjq(k, 4) Or dataW(i, 5) = dataJZjq(k, 5)) And (dataW(i, 1) = dataJZjq(k, 3) Or InStr(league, "预选") > 0) Then  '日期，主队，客队，联赛
                 Exit For
             End If
         Next
         
         If k <= UBound(dataJZjq) Then
             Call 记录初始值(dataSheet, dataJZjq, Loc, dataJZjqCol, k, 6, 6, True, False, "-", 8)
         End If
         
         '中国竞彩网半全场胜平负
         For k = 1 To UBound(dataJBqc)
             If dataW(i, 2) = dataJBqc(k, 1) And (dataW(i, 4) = dataJBqc(k, 4) Or dataW(i, 5) = dataJBqc(k, 5)) And (dataW(i, 1) = dataJBqc(k, 3) Or InStr(league, "预选") > 0) Then  '日期，主队，客队，联赛
                 Exit For
             End If
         Next
         
         If k <= UBound(dataJBqc) Then
             Call 记录初始值(dataSheet, dataJBqc, Loc, dataJBqcCol, k, 6, 6, True, False, "-", 9)
         End If
         
         '*******************************************************************
         '               竞彩网相关数据处理结束
         '********************************************************************
         
         'okbf1,okbf2  ------澳客网必发盈亏，okbf1 对应必发盈亏中的必发数据，okbf2对应必发盈亏中的99平均数据
         For k = 1 To UBound(dataBF)
             If dataW(i, 2) = dataBF(k, 2) And (dataW(i, 4) = dataBF(k, 4) Or dataW(i, 5) = dataBF(k, 5)) And dataW(i, 1) = dataBF(k, 1) Then '日期，对阵，联赛
                 Exit For
             End If
         Next
         
         If k <= UBound(dataBF) Then
             Call 记录初始值(dataSheet, dataBF, Loc, OKBF1col, k, 31, 31, True, True, "D", 3)
             Call 记录初始值(dataSheet, dataBF, Loc, OKBF2col, k, 37, 37, True, True, "D", 3)

             '2015.9.13，add 32,33,34配置
             For k1 = 2 To UBound(dataConfig)
                If dataColDict.exists(dataConfig(k1, 2)) And dataConfig(k1, 15) = "Y" Then
                    dataBgCol = dataColDict.Item(dataConfig(k1, 2))
                    initialBgCol = dataConfig(k1, 4)
                    realBgCol = dataConfig(k1, 5)
                    Application.Run dataConfig(k1, 11), dataSheet, dataBF, Loc, dataBgCol, k, initialBgCol, realBgCol, CBool(dataConfig(k1, 6)), CBool(dataConfig(k1, 7)), CStr(dataConfig(k1, 8)), CInt(dataConfig(k1, 9))
                    
                End If
             Next
         End If
         
         'Ok30 ---澳客网胜负数据
         For k = 1 To UBound(dataSF)
             If dataW(i, 2) = dataSF(k, 2) And (dataW(i, 4) = dataSF(k, 4) Or dataW(i, 5) = dataSF(k, 5)) And dataW(i, 1) = dataSF(k, 1) Then '日期，对阵，联赛
                 Exit For
             End If
         Next
         
         If k <= UBound(dataSF) Then
             Call 记录初始值(dataSheet, dataSF, Loc, OK30col, k, 7, 10, True, True, "A", 3)
         End If
         
         '-----------------------------
         '澳客网凯利方差
         '-----------------------------
         
        
        For k = 1 To UBound(dataKL)
            If dataW(i, 2) = dataKL(k, 2) And (dataW(i, 4) = dataKL(k, 4) Or dataW(i, 5) = dataKL(k, 5)) And dataW(i, 1) = dataKL(k, 1) Then  '日期，对阵，联赛
                Exit For
            End If
        Next
        
        If k <= UBound(dataKL) Then
            Call 记录初始值(dataSheet, dataKL, Loc, BFWcol, k, 7, 7, True, True, "D", 4)
            Call 记录初始值(dataSheet, dataKL, Loc, BFBcol, k, 11, 11, True, True, "D", 4)
            Call 记录初始值(dataSheet, dataKL, Loc, BFMcol, k, 15, 15, True, True, "D", 4)
            Call 记录初始值(dataSheet, dataKL, Loc, varCol, k, 19, 19, True, True, "D", 3)
        End If
          
          '将新增记录加入字典
        'dataDict.Add dataW(i, 0), loc
        If dataDict.exists(dataW(i, 0)) Then
            dataDict.Item(dataW(i, 0)) = Loc
        Else
            dataDict.Add dataW(i, 0), Loc
        End If
          
          
          '移动指针
          loc1 = loc1 + 3
          
      End If
  Next
  
  
  '----------------------------------------------------------
  '第三步：处理球探网其他类别的数据，采用字典的方式进行
  '----------------------------------------------------------
  
  
    '根据新的数据指针更新数据字典中的项值， add by ljqu 2016-5-18
    '将ID号和对应的行号载入字典,8列：球赛Id，0列：对应的excel行号
    If Not 载入综合数据字典(dataDict, data, 8, 0, "初始值") Then
        MsgBox ("更新综合数据指针字典时出错！")
        Exit Sub
    End If
      
  
  '-------处理球探网数据------------------
  
 'Bet365
 For k = 1 To UBound(dataB)
     vsId = dataB(k, 0)
     If dataDict.exists(vsId) Then      '根据赛事ID号来进行匹配
         Loc = dataDict.Item(vsId)
         If dataSheet.Cells(Loc, dataBcol) = "" And dataSheet.Cells(Loc, dataBcol + 1) = "" And dataSheet.Cells(Loc, dataBcol + 2) = "" Then
            Call 记录初始值(dataSheet, dataB, Loc, dataBcol, k, 7, 11, True, False, "A", 4)
         End If
     End If
 Next

 
 '澳门彩票
 For k = 1 To UBound(dataM)
     vsId = dataM(k, 0)
     If dataDict.exists(vsId) Then      '根据赛事ID号来进行匹配
         Loc = dataDict.Item(vsId)
         If dataSheet.Cells(Loc, dataMcol) = "" And dataSheet.Cells(Loc, dataMcol + 1) = "" And dataSheet.Cells(Loc, dataMcol + 2) = "" Then
            Call 记录初始值(dataSheet, dataM, Loc, dataMcol, k, 7, 11, True, False, "A", 4)
        End If
     End If
 Next


 '立博(英国)
 For k = 1 To UBound(dataL)
     vsId = dataL(k, 0)
     If dataDict.exists(vsId) Then      '根据赛事ID号来进行匹配
        Loc = dataDict.Item(vsId)
        If dataSheet.Cells(Loc, dataLcol) = "" And dataSheet.Cells(Loc, dataLcol + 1) = "" And dataSheet.Cells(Loc, dataLcol + 2) = "" Then
            Call 记录初始值(dataSheet, dataL, Loc, dataLcol, k, 7, 11, True, False, "A", 4)
        End If
     End If
 Next

 
 '易胜博
 For k = 1 To UBound(dataE)
     vsId = dataE(k, 0)
     If dataDict.exists(vsId) Then      '根据赛事ID号来进行匹配
         Loc = dataDict.Item(vsId)
         If dataSheet.Cells(Loc, dataEcol) = "" And dataSheet.Cells(Loc, dataEcol + 1) = "" And dataSheet.Cells(Loc, dataEcol + 2) = "" Then
            Call 记录初始值(dataSheet, dataE, Loc, dataEcol, k, 7, 11, True, False, "A", 4)
         End If
     End If
 Next
 
 
 
    '处理2020.03.15增加的数据
    For k1 = 2 To UBound(dataConfig)
       If dataColDict.exists(dataConfig(k1, 2)) And dataConfig(k1, 15) = "Y" And dataConfig(k1, 16) = "20200315" Then
           datasrc = dataConfig(k1, 10) '配置定义的数据源
           Select Case datasrc
           Case "dataOZ"
              datatmp = dataOZ
           Case "dataYZ"
              datatmp = dataYZ
           End Select
           
           If dataColDict.exists(dataConfig(k1, 2)) Then     '如果存在增中的数据列
                dataBgCol = dataColDict.Item(dataConfig(k1, 2))    '根据数据列标识，找到数据待插的首列
                initialBgCol = dataConfig(k1, 4)      '初始值对应的数据列
                realBgCol = dataConfig(k1, 5)         '实时值对应的数据列
                For k = 1 To UBound(datatmp)          '对加载的数据进行循环
                     vsId = datatmp(k, 0)
                     If dataDict.exists(vsId) Then      '根据赛事ID号来进行匹配
                        Loc = dataDict.Item(vsId)
                        Application.Run dataConfig(k1, 11), dataSheet, datatmp, Loc, dataBgCol, k, initialBgCol, realBgCol, CBool(dataConfig(k1, 6)), CBool(dataConfig(k1, 7)), CStr(dataConfig(k1, 8)), CInt(dataConfig(k1, 9))
                     End If
                Next
            End If
       End If
    Next



 '平均赔率1、平均赔率2、BF1
 For k = 1 To UBound(dataL12)
     vsId = dataL12(k, 0)
     If dataDict.exists(vsId) Then      '根据赛事ID号来进行匹配
         Loc = dataDict.Item(vsId)
         If dataSheet.Cells(Loc, LOSE1col) = "" And dataSheet.Cells(Loc, LOSE1col + 1) = "" And dataSheet.Cells(Loc, LOSE1col + 2) = "" Then
            Call 记录初始值(dataSheet, dataL12, Loc, LOSE1col, k, 1, 5, False, False, "A", 4) '平均赔率1
         End If
         If dataSheet.Cells(Loc, lose2Col) = "" And dataSheet.Cells(Loc, lose2Col + 1) = "" And dataSheet.Cells(Loc, lose2Col + 2) = "" Then
            Call 记录初始值(dataSheet, dataL12, Loc, lose2Col, k, 9, 13, False, False, "A", 4) '平均赔率2
         End If
         If dataSheet.Cells(Loc, BF1Col) = "" And dataSheet.Cells(Loc, BF1Col + 1) = "" And dataSheet.Cells(Loc, BF1Col + 2) = "" Then
            Call 记录初始值(dataSheet, dataL12, Loc, BF1Col, k, 25, 29, True, True, "D", 4)  'BF1
         End If
     End If
 Next
 
 
 '欧亚转盘赔率数据
 For k = 1 To UBound(dataJF)
    vsId = dataJF(k, 0)
    If dataDict.exists(vsId) Then
      
       Loc = dataDict.Item(vsId)
      'bet365 欧亚转盘 赔率
      If dataColDict.exists("E2AB") Then
          dataBgCol = dataColDict.Item("E2AB")
          If dataSheet.Cells(Loc, dataBgCol) = "" And dataSheet.Cells(Loc, dataBgCol + 1) = "" And dataSheet.Cells(Loc, dataBgCol + 2) = "" Then
            Call 记录初始值(dataSheet, dataJF, Loc, dataBgCol, k, 89, 93, True, False, "false", 4)
          End If
      End If
      
      '澳彩   欧亚转盘 赔率
      If dataColDict.exists("E2AM") Then
          dataBgCol = dataColDict.Item("E2AM")
          If dataSheet.Cells(Loc, dataBgCol) = "" And dataSheet.Cells(Loc, dataBgCol + 1) = "" And dataSheet.Cells(Loc, dataBgCol + 2) = "" Then
            Call 记录初始值(dataSheet, dataJF, Loc, dataBgCol, k, 97, 101, True, False, "false", 4)
          End If
      End If
    End If
Next
 
    
  MsgBox ("初始数据完毕！")
  Set dataSheet = Nothing
End Sub


Sub 数据更新(ByRef control As Office.IRibbonControl)

    Call 数据更新实现

End Sub

Sub 数据更新实现(Optional hisFlag As Boolean = False)
'******************************************************************************************
'
'即时值更新，前一个即时值移入即时值1，新进的数据写入到即时值2
'
'******************************************************************************************

Dim wkWorkbook As Workbook

Dim wkSheet As Worksheet
Dim dataSheet As Worksheet

Dim data()
Dim dataW()
Dim dataB()
Dim dataM()
Dim dataL()
Dim dataE()
Dim dataJ()
Dim dataL12()   '赔1、赔2
Dim dataJF()    '球探网联赛积分
Dim dataOZ()    '欧指数据   2020.03.15
Dim dataYZ()    '亚指数据   2020.03.15
Dim datatmp()   '临时数据存放 2020.03.15
Dim datasrc As String  '临时数据存放 2020.03.15

Dim dataBF()   '澳客网必发
Dim dataSF()   '澳客网胜负
Dim dataKL()   '澳客网凯利
Dim dataOK()   '澳客网期数
Dim dataPK()   '澳客网盘口预测

Dim dataConfig()   '配置数据信息


Dim dataJBf()     '竞彩网比分
Dim dataJZjq()    '竞彩网总进球
Dim dataJBqc()    '竞彩网半全场胜平负

Dim i As Long
Dim j As Long
Dim k As Long
Dim k1 As Long
Dim Loc As Long
Dim loc1 As Long
Dim wloc As Long      '威廉数据的记录个数
Dim tloc As Long      '总的数据循环次数


'定义赛事信息
Dim league        '联赛
Dim currdate      '赛事日期
Dim vsInfo        '对阵信息
Dim priTeam       '主队
Dim secTeam       '客队
Dim vsId          '球探网赛事ID

'参数

Dim calMeth As String    '数据处理方式，1：按最小记录进行处理，9：按最大记录处理
Dim dataWcol As Integer   '威廉希尔数据开始列号
Dim dataBcol As Integer     'Bet365数据开始列号
Dim dataMcol As Integer     '澳门彩票数据开始列号
Dim dataLcol As Integer     '立博(英国)数据开始列号
Dim dataEcol As Integer     '易胜博数据开始列号
Dim DataJcol As Integer     '竞彩网数据开始列号
Dim LOSE1col As Integer     '赔1 数据开始列号
Dim lose2Col As Integer     '赔2 数据开始列号
Dim BF1Col   As Integer     'BF1 数据开始列号
Dim BF2col  As Integer      'BF2 数据开始列号
Dim BF3col  As Integer      'BF3 数据开始列号
Dim OKBF1col As Integer     'OKBF1数据开始列号
Dim OKBF2col As Integer     'OKBF2数据开始列号
Dim OKnetCurrentCol As Integer   '澳客网当期期数存放列
Dim OK30col As Integer      'Ok30数据——澳客网胜负数据
Dim BFWcol As Integer       '威廉希尔——澳客网凯利方差
Dim BFMcol As Integer       'Bet365——澳客网凯利方差
Dim BFBcol As Integer       '澳门——澳客网凯利方差
Dim varCol As Integer       '方差——澳客网凯利方差

Dim varCmpCol As Long       '方差分析数据开始列号
Dim dataJRqCol As Long      '竞彩网让球数据开始列号
Dim dataJBfCol As Long      '竞彩网比分数据开始列号
Dim dataJZjqCol As Long     '竞彩网总让球开始列号
Dim dataJBqcCol As Long     '竞彩网半全场胜平负数据开始列号

Dim resultCol As Long       '比分数据开始列
Dim winloseCol As Long      '比赛结果数据开始列
Dim turnsCol As Long        '轮次数据开始列
Dim rqcntCol As Long        '让球个数数据开始列

Dim usrCol As Integer       '用户临时数据列



Dim dataBgCol As Long    '数据在SHEET对应的数据起始列
Dim initialBgCol As Integer   '初始数据在数组中的起始列
Dim realBgCol As Integer   '即时数据在数组中的起始列
Dim funName As String      '动态调用的函数


Dim isCollectOknet As Boolean   '是否采集澳客网数据
Dim isFind As Boolean           '是否找到相应的记录
Dim dealRecCount As Long        '综合数据加载到内存的记录条数

Dim cnt

Dim insertRecno As Long         '新记录插入的位置信息

Dim tmClDict As Object    '联赛对应分类字典
Dim clStrDict As Object   '排名对应积分字典

Dim ll, panmBgCol


Call 初始化字典(Dict, "Param")
Call 初始化字典(leagueDict, "01赛事")
Call 初始化字典(tmClDict, "01赛事", 2, 1, 6)
Call 初始化字典(clStrDict, "TeamClass", 3, 1, 13)
Call 配置数据载入(dataConfig, "Config")

calMeth = Dict.Item("CALMETH")





dealRecCount = CLng(Dict.Item("DEAL_RECCOUNT"))
dataBgCol = CLng(Dict.Item("DATABGCOL"))

isCollectOknet = CBool(Dict.Item("IS_COLLECT_OKNET"))



'将综合数据加载入内存

If hisFlag Then
    Call 指定日期综合数据载入(data)
Else
    Call 综合数据载入内存(data, "综合数据", dealRecCount, dataBgCol)
End If
tloc = UBound(data)


Set dataSheet = ActiveWorkbook.Sheets("综合数据")
Call 初始化一般字典(dataColDict, dataSheet, 4, 0, 1, False)

'判断最近升级是否完毕，2017.10.15增加
usrCol = dataColDict.Item("SCHEMA6_1")
If usrCol = 0 Then
    'MsgBox ("程序升级中，请勿操作EXCEL......")
    'Call 程序升级
    MsgBox ("请先升级程序，再执行更新！")
    Exit Sub
End If


dataWcol = dataColDict.Item("DATAW")
dataBcol = dataColDict.Item("DATAB")
dataMcol = dataColDict.Item("DATAM")
dataLcol = dataColDict.Item("DATAL")
dataEcol = dataColDict.Item("DATAE")
DataJcol = dataColDict.Item("DATAJ")

LOSE1col = dataColDict.Item("LOSE1")
lose2Col = dataColDict.Item("LOSE2")
BF1Col = dataColDict.Item("BF1")
BF2col = dataColDict.Item("BF2")
BF3col = dataColDict.Item("BF3")

OKBF1col = dataColDict.Item("OKBF1")
OKBF2col = dataColDict.Item("OKBF2")

OK30col = dataColDict.Item("OK30")
BFWcol = dataColDict.Item("BFW")
BFMcol = dataColDict.Item("BFM")
BFBcol = dataColDict.Item("BFB")
varCol = dataColDict.Item("VAR")


varCmpCol = dataColDict.Item("VARCMP")
dataJRqCol = dataColDict.Item("DATAJRQ")
dataJBfCol = dataColDict.Item("DATAJBF")
dataJZjqCol = dataColDict.Item("DATAJZJQ")
dataJBqcCol = dataColDict.Item("DATAJBQC")

resultCol = dataColDict.Item("RESULT")
winloseCol = dataColDict.Item("WINLOSE")
turnsCol = dataColDict.Item("TURNS")
rqcntCol = dataColDict.Item("RQCNT")

OKnetCurrentCol = dataColDict.Item("OKID")

'一
Call LoadDataToArray(dataW, "球探网(W)")
wloc = UBound(dataW)

'依据威廉希尔的数据去追踪bf3数据、赔1数据、赔2数据、以及bf3的数据
Call BF数据载入(dataL12, "球探网(BF)")

'--------------------------------------------
'  从2014.10.6开始，B、M、L、E的数据从BF中取
'--------------------------------------------
Call BF载入赔率(dataB, "球探网(BF)", "B")      'bet365(英国)
Call BF载入赔率(dataM, "球探网(BF)", "M")       '澳门
Call BF载入赔率(dataL, "球探网(BF)", "L")       '立博（英国）
Call BF载入赔率(dataE, "球探网(BF)", "E")       '易胜博
Call 球探网联赛积分载入(dataJF, "球探网积分")    '联赛积分
Call 欧指数据载入(dataOZ, "OZ")                 '欧指数据 2020.03.15
Call 亚指数据载入(dataYZ, "YZ")                 '亚指数据 2020.03.15

'-------------------------------------------
'    澳客网数据载入
'-------------------------------------------
Call 澳客网必发盈亏载入(dataBF, "澳客网(1)")
Call 澳客网凯利方差载入(dataKL, "澳客网(3)")
Call 澳客网胜负指数载入(dataSF, "澳客网(2)")
Call 澳客网胜负指数载入(dataOK, "澳客网期数")
Call 澳客网盘口评测载入(dataPK, "澳客网(4)")

Call 加载中国竞彩网数据(dataJ, "中国竞彩网")

Call 加载中国竞彩网数据(dataJ, "中国竞彩网")
Call 加载竞彩网比分数据(dataJBf, "竞彩网比分")
Call 加载竞彩网总进球数据(dataJZjq, "竞彩网总进球")
Call 加载竞彩网半全场胜平负数据(dataJBqc, "竞彩网半全场")

j = 1
Do While j <= tloc
    '移动指针
    If data(j, 6) = "即时值2" Then
        league = data(j, 7)       '联赛
        currdate = data(j, 1)   '赛事日期
        vsInfo = data(j, 5)        '对阵信息
        priTeam = data(j, 3)      '主队
        secTeam = data(j, 4)     '客队
        vsId = data(j, 8)        '球探网赛事ID
        
        Loc = data(j, 0)       '获取数据在sheet中的编号
        
        
         '威廉希尔
        If UBound(dataW) > 0 Then
            If currdate >= dataW(1, 2) Then     '小于最起始的时间直接跳过
                For k = 1 To UBound(dataW)
                    'If currDate = dataW(k, 2) And vsInfo = dataW(k, 6) And league = dataW(k, 1) Then '日期，对阵，联赛
                    If vsId = dataW(k, 0) Then   '直接比较球赛ID号
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataW) Then
                    Call 记录即时值(dataSheet, dataW, Loc, dataWcol, k, 7, 11, False, "A", 4)
                    If dataSheet.Cells(Loc - 2, 10) = "" And dataW(k, 15) <> "" Then
                        dataSheet.Cells(Loc - 2, 10) = dataW(k, 15)
                        dataSheet.Cells(Loc - 2, 11) = 比赛结果(dataW(k, 15))
                        
                        dataSheet.Cells(Loc - 1, 10) = dataSheet.Cells(Loc - 2, 10) '比分
                        dataSheet.Cells(Loc - 1, 11) = dataSheet.Cells(Loc - 2, 11) '赛果
                        dataSheet.Cells(Loc, 10) = dataSheet.Cells(Loc - 2, 10) '比分
                        dataSheet.Cells(Loc, 11) = dataSheet.Cells(Loc - 2, 11) '赛果
                    End If
                End If
            End If
        End If
        
        'Bet365
        If UBound(dataB) > 0 Then
             If currdate >= dataB(1, 2) Then     '小于最起始的时间直接跳过
                For k = 1 To UBound(dataB)
                    'If currDate = dataB(k, 2) And vsInfo = dataB(k, 6) And league = dataB(k, 1) Then '日期，对阵，联赛
                    If vsId = dataB(k, 0) Then   '直接比较球赛ID号
                        Exit For
                    End If
                Next
    
                If k <= UBound(dataB) Then
                    Call 记录即时值(dataSheet, dataB, Loc, dataBcol, k, 7, 11, False, "A", 4)
                End If
            End If
        End If
        
        '澳门彩票
        If UBound(dataM) > 0 Then
            If currdate >= dataM(1, 2) Then     '小于最起始的时间直接跳过
                For k = 1 To UBound(dataM)
                    'If currDate = dataM(k, 2) And vsInfo = dataM(k, 6) And league = dataM(k, 1) Then '日期，对阵，联赛
                    If vsId = dataM(k, 0) Then   '直接比较球赛ID号
                        Exit For
                    End If
                Next
                
                 If k <= UBound(dataM) Then
                    Call 记录即时值(dataSheet, dataM, Loc, dataMcol, k, 7, 11, False, "A", 4)
                End If
            End If
        End If

        
        '立博(英国)
        If UBound(dataL) > 0 Then
            If currdate >= dataL(1, 2) Then     '小于最起始的时间直接跳过
                For k = 1 To UBound(dataL)
                    'If currDate = dataL(k, 2) And vsInfo = dataL(k, 6) And league = dataL(k, 1) Then '日期，对阵，联赛
                    If vsId = dataL(k, 0) Then   '直接比较球赛ID号
                        Exit For
                    End If
                Next
                
                 If k <= UBound(dataL) Then
                    Call 记录即时值(dataSheet, dataL, Loc, dataLcol, k, 7, 11, False, "A", 4)
                End If
            End If
        End If
        
        
        '易胜博
        If UBound(dataE) > 0 Then
            If currdate >= dataE(1, 2) Then     '小于最起始的时间直接跳过
                For k = 1 To UBound(dataE)
                    'If currDate = dataE(k, 2) And vsInfo = dataE(k, 6) And league = dataE(k, 1) Then '日期，对阵，联赛
                    If vsId = dataE(k, 0) Then   '直接比较球赛ID号
                        Exit For
                    End If
                Next
                
                 If k <= UBound(dataE) Then
                    Call 记录即时值(dataSheet, dataE, Loc, dataEcol, k, 7, 11, False, "A", 4)
                End If
            
            End If
        End If
        
        
        '处理2020.03.15增加的数据
        For k1 = 2 To UBound(dataConfig)
           If dataColDict.exists(dataConfig(k1, 2)) And dataConfig(k1, 15) = "Y" And dataConfig(k1, 16) = "20200315" Then
               datasrc = dataConfig(k1, 10) '配置定义的数据源
               Select Case datasrc
               Case "dataOZ"
                  datatmp = dataOZ
               Case "dataYZ"
                  datatmp = dataYZ
               End Select
               
               If UBound(datatmp) > 0 Then
                    dataBgCol = dataColDict.Item(dataConfig(k1, 2))    '根据数据列标识，找到数据待插的首列
                    initialBgCol = dataConfig(k1, 4)      '初始值对应的数据列
                    realBgCol = dataConfig(k1, 5)         '实时值对应的数据列
                    For k = 1 To UBound(datatmp)          '对加载的数据进行循环
                         If vsId = datatmp(k, 0) Then   '直接比较球赛ID号
                            Application.Run dataConfig(k1, 12), dataSheet, datatmp, Loc, dataBgCol, k, initialBgCol, realBgCol, CBool(dataConfig(k1, 7)), CStr(dataConfig(k1, 8)), CInt(dataConfig(k1, 9))
                         End If
                    Next
                End If
           End If
        Next
        
        '2020.03.15计算必发指数
        Call 计算必发指数(dataColDict, dataSheet, Loc)
        
        '**************************************************************************************
        '                      中国竞彩网数据处理
        '**************************************************************************************
        
        '中国竞彩网
        If UBound(dataJ) > 0 Then
            If currdate >= dataJ(1, 1) Then    '小于最起始的时间直接跳过
                For k = 1 To UBound(dataJ)
                    If currdate = dataJ(k, 1) And (priTeam = dataJ(k, 4) Or secTeam = dataJ(k, 5)) And (league = dataJ(k, 3) Or InStr(league, "预选") > 0) Then '日期，主队，客队，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataJ) Then
                    Call 记录即时值(dataSheet, dataJ, Loc, DataJcol, k, 6, 9, True, "D", 3)
                    Call 记录即时值(dataSheet, dataJ, Loc, dataJRqCol, k, 19, 19, True, "D", 3)    '让球投资比例
                    '让球个数
                    dataSheet.Cells(Loc, rqcntCol) = dataJ(k, 12)
                    dataSheet.Cells(Loc - 1, rqcntCol) = dataJ(k, 12)
                    dataSheet.Cells(Loc - 2, rqcntCol) = dataJ(k, 12)
                    
                End If
            End If
        End If
        
        
        '中国竞彩网比分
        If UBound(dataJBf) > 0 Then
            If currdate >= dataJBf(1, 1) Then    '小于最起始的时间直接跳过
                For k = 1 To UBound(dataJBf)
                    If currdate = dataJBf(k, 1) And (priTeam = dataJBf(k, 4) Or secTeam = dataJBf(k, 5)) And (league = dataJBf(k, 3) Or InStr(league, "预选") > 0) Then '日期，主队，客队，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataJBf) Then
                    Call 记录即时值(dataSheet, dataJBf, Loc, dataJBfCol, k, 6, 6, False, "-", 31)
                End If
            End If
        End If
        
        
        '中国竞彩网总进球
        If UBound(dataJZjq) > 0 Then
            If currdate >= dataJZjq(1, 1) Then    '小于最起始的时间直接跳过
                For k = 1 To UBound(dataJZjq)
                    If currdate = dataJZjq(k, 1) And (priTeam = dataJZjq(k, 4) Or secTeam = dataJZjq(k, 5)) And (league = dataJZjq(k, 3) Or InStr(league, "预选") > 0) Then '日期，主队，客队，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataJZjq) Then
                    Call 记录即时值(dataSheet, dataJZjq, Loc, dataJZjqCol, k, 6, 6, False, "-", 8)
                End If
            End If
        End If
        
        
        '中国竞彩网半全场胜平负
        If UBound(dataJBqc) > 0 Then
            If currdate >= dataJBqc(1, 1) Then    '小于最起始的时间直接跳过
                For k = 1 To UBound(dataJBqc)
                    If currdate = dataJBqc(k, 1) And (priTeam = dataJBqc(k, 4) Or secTeam = dataJBqc(k, 5)) And (league = dataJBqc(k, 3) Or InStr(league, "预选") > 0) Then '日期，主队，客队，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataJBqc) Then
                    Call 记录即时值(dataSheet, dataJBqc, Loc, dataJBqcCol, k, 6, 6, False, "-", 9)
                End If
            End If
        End If
        
        '**************************************************************************************
        '                      中国竞彩网数据处理结束
        '**************************************************************************************
        
        If UBound(dataW) > 0 Then
             If currdate >= dataW(1, 2) Then     '小于最起始的时间直接跳过,由于此数据依附于dataW
                  '平均赔率1
                  For k = 1 To UBound(dataL12)
                      If vsId = dataL12(k, 0) Then 'ID相等
                          Exit For
                      End If
                  Next
                  
                  If k <= UBound(dataL12) Then
                      Call 记录即时值(dataSheet, dataL12, Loc, LOSE1col, k, 1, 5, False, "A", 4)
                  End If
        
                  
                  '平均赔率2
                  For k = 1 To UBound(dataL12)
                      If vsId = dataL12(k, 0) Then 'ID相等
                          Exit For
                      End If
                  Next
                  
                  If k <= UBound(dataL12) Then
                      Call 记录即时值(dataSheet, dataL12, Loc, lose2Col, k, 9, 13, False, "A", 4)
                  End If
                  
                  
                  'BF1  --有四个数据
                  For k = 1 To UBound(dataL12)
                      If vsId = dataL12(k, 0) Then 'ID相等
                          Exit For
                      End If
                  Next
                  
                  If k <= UBound(dataL12) Then
                      Call 记录即时值(dataSheet, dataL12, Loc, BF1Col, k, 25, 29, True, "D", 4)
                  End If
                  
                  '------------------------------------
                  '球探网联赛职分    add 2015.3.29
                  '------------------------------------
                  For k = 1 To UBound(dataJF)
                      If vsId = dataJF(k, 0) Then 'ID相等
                          Exit For
                      End If
                  Next
                  
                  If k <= UBound(dataJF) Then
                        '主队积分
                        If dataColDict.exists("SCOREM") Then
                            dataBgCol = dataColDict.Item("SCOREM")
                            Call 记录联赛积分即时值(dataSheet, dataJF, Loc, dataBgCol, k, 3, 14, 36, False, "false", 5)
                        End If
                        
                        '客队积分
                        If dataColDict.exists("SCORES") Then
                            dataBgCol = dataColDict.Item("SCORES")
                            Call 记录联赛积分即时值(dataSheet, dataJF, Loc, dataBgCol, k, 47, 69, 80, False, "false", 5)
                        End If
                        
                        'bet365 欧亚转盘 赔率
                        If dataColDict.exists("E2AB") Then
                            dataBgCol = dataColDict.Item("E2AB")
                            Call 记录即时值(dataSheet, dataJF, Loc, dataBgCol, k, 89, 93, False, "false", 4)
                        End If
                        
                        Call 计算盘形分析值(dataSheet, Loc)
                        
                        '澳彩   欧亚转盘 赔率
                        If dataColDict.exists("E2AM") Then
                            dataBgCol = dataColDict.Item("E2AM")
                            Call 记录即时值(dataSheet, dataJF, Loc, dataBgCol, k, 97, 101, False, "false", 4)
                        End If
                  End If
                  
                  
            End If
        End If
        
        
        If UBound(dataKL) > 0 Then
            If currdate >= dataKL(1, 2) Then     '小于最起始的时间直接跳过
            
                'OkBf1、Okbf2 ---澳客网必发盈亏
                For k = 1 To UBound(dataBF)
                    If currdate = dataBF(k, 2) And (InStr(priTeam, dataBF(k, 4)) > 0 Or InStr(secTeam, dataBF(k, 5)) > 0) And league = dataBF(k, 1) Then '日期，对阵，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataBF) Then
                    Call 记录即时值(dataSheet, dataBF, Loc, OKBF1col, k, 31, 31, True, "D", 3)
                    Call 记录即时值(dataSheet, dataBF, Loc, OKBF2col, k, 37, 37, True, "D", 3)
                                        
                    '2015.9.13，add 32,33,34配置
                    '2016.5.18, 在配置中删除了三项，因此位置由21-34，调整为18-31。
                    '2016.8.12, 在config页中增加“启用标志”列
                    For k1 = 2 To UBound(dataConfig)
                        If dataColDict.exists(dataConfig(k1, 2)) And dataConfig(k1, 15) = "Y" And dataConfig(k1, 16) = "" Then
                            dataBgCol = dataColDict.Item(dataConfig(k1, 2))
                            initialBgCol = dataConfig(k1, 4)
                            realBgCol = dataConfig(k1, 5)
                            Application.Run dataConfig(k1, 12), dataSheet, dataBF, Loc, dataBgCol, k, initialBgCol, realBgCol, CBool(dataConfig(k1, 7)), CStr(dataConfig(k1, 8)), CInt(dataConfig(k1, 9))
                        End If
                    Next
                End If
            
            
                'Ok30 ---澳客网胜负数据
                For k = 1 To UBound(dataSF)
                    If currdate = dataSF(k, 2) And (InStr(priTeam, dataSF(k, 4)) > 0 Or InStr(secTeam, dataSF(k, 5)) > 0) And league = dataSF(k, 1) Then '日期，对阵，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataSF) Then
                    Call 记录即时值(dataSheet, dataSF, Loc, OK30col, k, 7, 10, True, "A", 3)
                End If
                
                
                 '-----------------------------
                '澳客网凯利方差
                '-----------------------------
                
    
                For k = 1 To UBound(dataKL)
                    If currdate = dataKL(k, 2) And (InStr(priTeam, dataKL(k, 4)) > 0 Or InStr(secTeam, dataKL(k, 5)) > 0) And league = dataKL(k, 1) Then  '日期，对阵，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataKL) Then
                    Call 记录即时值(dataSheet, dataKL, Loc, BFWcol, k, 7, 7, True, "D", 4)
                    Call 记录即时值(dataSheet, dataKL, Loc, BFBcol, k, 11, 11, True, "D", 4)
                    Call 记录即时值(dataSheet, dataKL, Loc, BFMcol, k, 15, 15, True, "D", 4)
                    Call 记录即时值(dataSheet, dataKL, Loc, varCol, k, 19, 19, True, "D", 3)
                    If dataSheet.Cells(Loc - 2, resultCol) = "" And "VS" <> UCase(dataKL(k, 6)) Then
                        dataSheet.Cells(Loc - 2, resultCol) = dataKL(k, 6)
                        dataSheet.Cells(Loc - 1, resultCol) = dataKL(k, 6)
                        dataSheet.Cells(Loc, resultCol) = dataKL(k, 6)
                        dataSheet.Cells(Loc - 2, winloseCol) = 比赛结果(dataKL(k, 6))
                        dataSheet.Cells(Loc - 1, winloseCol) = dataSheet.Cells(Loc - 2, winloseCol)
                        dataSheet.Cells(Loc, winloseCol) = dataSheet.Cells(Loc - 2, winloseCol)
                    End If
                    
                    
                     'add by ljqu 2016.8.12   增加方差分析
                    Call 方差分析及记录(dataSheet, Loc, varCmpCol, varCol, 0, True, "D", 3)
                    
                End If
                
                
               
                
                
                
                '-------------------------------
                '澳客网盘口评测
                '-------------------------------
                For k = 1 To UBound(dataPK)
                    If currdate = dataPK(k, 2) And (InStr(priTeam, dataPK(k, 4)) > 0 Or InStr(secTeam, dataPK(k, 5)) > 0) And league = dataPK(k, 1) Then '日期，对阵，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataPK) Then
                    '澳门盘口数据
                    If dataColDict.exists("PANM") Then
                        dataBgCol = dataColDict.Item("PANM")
                        Call 记录盘口即时值(dataSheet, dataPK, Loc, dataBgCol, k, 30, 46, False, "False", 4)
                        For ll = 1 To 3
                            panmBgCol = dataColDict.Item("PANM_" & ll)
                            dataSheet.Cells(Loc - 2, panmBgCol) = dataSheet.Cells(Loc - 3 + ll, dataBgCol)
                            dataSheet.Cells(Loc - 2, panmBgCol + 1) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 1)
                            dataSheet.Cells(Loc - 2, panmBgCol + 2) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 2)
                            
                            dataSheet.Cells(Loc - 1, panmBgCol) = dataSheet.Cells(Loc - 3 + ll, dataBgCol)
                            dataSheet.Cells(Loc - 1, panmBgCol + 1) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 1)
                            dataSheet.Cells(Loc - 1, panmBgCol + 2) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 2)
                            
                            dataSheet.Cells(Loc, panmBgCol) = dataSheet.Cells(Loc - 3 + ll, dataBgCol)
                            dataSheet.Cells(Loc, panmBgCol + 1) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 1)
                            dataSheet.Cells(Loc, panmBgCol + 2) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 2)
                        Next
                    End If
                    'Bet365盘口数据
                    If dataColDict.exists("PANB") Then
                        dataBgCol = dataColDict.Item("PANB")
                        Call 记录盘口即时值(dataSheet, dataPK, Loc, dataBgCol, k, 10, 26, False, "False", 4)
                        For ll = 1 To 3
                            panmBgCol = dataColDict.Item("PANB_" & ll)
                            dataSheet.Cells(Loc - 2, panmBgCol) = dataSheet.Cells(Loc - 3 + ll, dataBgCol)
                            dataSheet.Cells(Loc - 2, panmBgCol + 1) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 1)
                            dataSheet.Cells(Loc - 2, panmBgCol + 2) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 2)
                            
                            dataSheet.Cells(Loc - 1, panmBgCol) = dataSheet.Cells(Loc - 3 + ll, dataBgCol)
                            dataSheet.Cells(Loc - 1, panmBgCol + 1) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 1)
                            dataSheet.Cells(Loc - 1, panmBgCol + 2) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 2)
                            
                            dataSheet.Cells(Loc, panmBgCol) = dataSheet.Cells(Loc - 3 + ll, dataBgCol)
                            dataSheet.Cells(Loc, panmBgCol + 1) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 1)
                            dataSheet.Cells(Loc, panmBgCol + 2) = dataSheet.Cells(Loc - 3 + ll, dataBgCol + 2)
                        Next
                    End If
                End If
                
                
                'Ok30 ---更新竞彩期数
                For k = 1 To UBound(dataOK)
                    If currdate = dataOK(k, 2) And (InStr(priTeam, dataOK(k, 4)) > 0 Or InStr(secTeam, dataOK(k, 5)) > 0) And league = dataOK(k, 1) Then '日期，对阵，联赛
                        Exit For
                    End If
                Next
                
                If k <= UBound(dataOK) Then
                    dataSheet.Cells(Loc - 2, OKnetCurrentCol) = dataOK(k, 13)
                    dataSheet.Cells(Loc - 2, OKnetCurrentCol + 1) = dataOK(k, 14)
                    dataSheet.Cells(Loc - 1, OKnetCurrentCol) = dataOK(k, 13)
                    dataSheet.Cells(Loc - 1, OKnetCurrentCol + 1) = dataOK(k, 14)
                    dataSheet.Cells(Loc, OKnetCurrentCol) = dataOK(k, 13)
                    dataSheet.Cells(Loc, OKnetCurrentCol + 1) = dataOK(k, 14)
                End If
            
            End If
        End If
        
        Call 排名分析(dataSheet, Loc, dataColDict, tmClDict, clStrDict, league)
        Call 主客队排名分析(dataSheet, Loc, dataColDict, tmClDict, clStrDict, league)
        
        If hisFlag Then
            j = j + 1
        Else
            j = j + 3
        End If
    Else
        j = j + 1
    End If
Loop

    MsgBox ("数据更新完毕！")
    Set dataSheet = Nothing
End Sub

Sub 模式计算(ByRef control As Office.IRibbonControl)
Dim data()
Dim configData()
Dim dataSheet As Worksheet
Dim cnt
Dim i, j, k

'跟参数配置相关的变量

Dim k1, j1        '参数配置处理中用至循环变量
Dim dataBeginCol As Integer     '数据开始的列号
Dim paraCnt As Integer       '基础数据长度
Dim paraCompType As String             '比较处理类型
Dim upColor
Dim downColor



Dim Loc As Long
Dim usrCol As Integer       '用户临时数据列

Dim dataWcol As Integer   '威廉希尔数据开始列号
Dim dataBcol As Integer     'Bet365数据开始列号
Dim dataMcol As Integer     '澳门彩票数据开始列号
Dim dataLcol As Integer     '立博(英国)数据开始列号
Dim dataEcol As Integer     '易胜博数据开始列号
Dim LOSE1col As Integer      '赔1数据列号
Dim lose2Col As Integer      '赔2数据列号
Dim BF1Col As Integer        'BF1数据开始列号
Dim BF2col As Integer        'BF2数据开始列号
Dim BF3col As Integer        'BF3数据开始列号
Dim varCol As Integer        '方差数据开始列号
Dim panmCol As Integer      '澳门盘口开始列号
Dim panbCol As Integer      'Bet365盘口开始列号

Dim recCnt  As Long     '待取的数据个数

Dim offset1 As Integer         '至第1个数据的偏移量
Dim offset2 As Integer         '至第2个数据的偏移量
Dim offset3 As Integer         '至第3个数据的偏移量
Dim offset4 As Integer         '至第4个数据的偏移量


Dim schemaCol1 As Integer
Dim schemaCol2 As Integer
Dim schemaCol3 As Integer

Dim schemaCol4 As Integer
Dim schemaCol5 As Integer
Dim schemaCol6 As Integer
Dim schemaCol7 As Integer
Dim schemaCol8 As Integer

Dim schemaCol11 As Integer
Dim schemaCol21 As Integer
Dim schemaCol41 As Integer

Dim schemaCol61 As Integer
Dim schemaCol71 As Integer
Dim schemaCol81 As Integer


Dim dataCol(6)


Dim value1 As Double
Dim value2 As Double
Dim value3 As Double


upColor = vbRed
downColor = vbGreen

Call 初始化字典(Dict, "Param")

recCnt = CLng(Dict.Item("DEAL_RECCOUNT"))


Call 配置数据载入(configData, "Config")

Set dataSheet = ActiveWorkbook.Sheets("综合数据")
cnt = dataSheet.UsedRange.Rows(dataSheet.UsedRange.Rows.Count).row
ReDim data(cnt - 4, 29) '对应worksheet中的行号，日期，联赛，对阵，数据类型，五组数据（每组3个)，模式四组（1-4），2018.8.20改为五组改为七组（加上赔1和赔2的数据）

Call 初始化一般字典(dataColDict, dataSheet, 4, 0, 1, False)



dataCol(0) = dataColDict.Item("DATAW")
dataCol(1) = dataColDict.Item("DATAB")
dataCol(2) = dataColDict.Item("DATAM")
dataCol(3) = dataColDict.Item("DATAL")
dataCol(4) = dataColDict.Item("DATAE")

dataCol(5) = dataColDict.Item("LOSE1")
dataCol(6) = dataColDict.Item("LOSE2")





LOSE1col = dataColDict.Item("LOSE1")
lose2Col = dataColDict.Item("LOSE2")
BF1Col = dataColDict.Item("BF1")
BF2col = dataColDict.Item("BF2")
BF3col = dataColDict.Item("BF3")
varCol = dataColDict.Item("VAR")
panmCol = dataColDict.Item("PANM")
panbCol = dataColDict.Item("PANB")

schemaCol1 = dataColDict.Item("SCHEMA1")
schemaCol2 = dataColDict.Item("SCHEMA2")
schemaCol3 = dataColDict.Item("SCHEMA3")
schemaCol4 = dataColDict.Item("SCHEMA4")   '模式四的起始列
schemaCol5 = dataColDict.Item("SCHEMA5")   '模式五的起始列
schemaCol6 = dataColDict.Item("SCHEMA6")   '模式六的起始列
schemaCol7 = dataColDict.Item("SCHEMA7")   '模式七的起始列
schemaCol8 = dataColDict.Item("SCHEMA8")   '模式八的起始列


schemaCol11 = dataColDict.Item("SCHEMA1_1")
schemaCol21 = dataColDict.Item("SCHEMA2_1")
schemaCol41 = dataColDict.Item("SCHEMA4_1")

schemaCol61 = dataColDict.Item("SCHEMA6_1")
schemaCol71 = dataColDict.Item("SCHEMA7_1")
schemaCol81 = dataColDict.Item("SCHEMA8_1")


'判断最近升级是否完毕
usrCol = dataColDict.Item("DATAB_1")
If usrCol = 0 Then
    MsgBox ("2018.8.27升级未完成！请先执行程序升级....")
    Exit Sub
End If


'将数据取到内存数组中

'确定取数的范围，为0则取全部
If recCnt = 0 Then
    i = 5
ElseIf cnt - recCnt + 1 < 4 Then
    i = 5
Else
    i = cnt - recCnt + 1
    ReDim data(recCnt - 1, 29) '对应worksheet中的行号，日期，联赛，对阵，数据类型七组数据（每组3个)，模式四组（1-4），2018.8.20改为五组改为七组（加上赔1和赔2的数据）
End If

Loc = 0
Do While i <= cnt
    If dataSheet.Cells(i, 1) <> "" Then   'And dataSheet.Cells(i, 3) <> "" And dataSheet.Cells(i, 5) <> "" Then
        'If dataSheet.Cells(i, 6) = "初始值" Then
            data(Loc, 0) = i      '对应workSheet中的行号，用于在回填时直接写到对应的单元格
            data(Loc, 1) = dataSheet.Cells(i, 1) '日期
            data(Loc, 2) = dataSheet.Cells(i, 7) '联赛
            data(Loc, 3) = dataSheet.Cells(i, 5)  '对阵
            data(Loc, 4) = dataSheet.Cells(i, 6)  '数据类型
            k = 5
            For j = 0 To 6
                data(Loc, k + 3 * j) = dataSheet.Cells(i, dataCol(j))
                data(Loc, k + 3 * j + 1) = dataSheet.Cells(i, dataCol(j) + 1)
                data(Loc, k + 3 * j + 2) = dataSheet.Cells(i, dataCol(j) + 2)
            Next
            '指针移位
            Loc = Loc + 1
        '    i = i + 3
        'Else
            i = i + 1
        'End If
    Else
        Exit Do
    End If
Loop

Call 模式计算一(data, "0,1,2", "W,B,M", 5, 26, 1)  '模式一
Call 模式计算一(data, "0,1,2", "3,1,0", 5, 27, 2)  '模式二
Call 模式计算一(data, "0,4,3", "W,E,L", 5, 28, 1)   '模式三
'Call 模式计算一(data, "0,1,2", "W,B,M", 5, 23)   '模式四


'求模式四的计算公式

offset1 = dataCol(0) - schemaCol4    '威廉希尔起始列-模式四起始列
offset2 = dataCol(1) - schemaCol4    'Bet365起始列-模式四起始列
offset3 = dataCol(2) - schemaCol4    '澳门网起始列-模式四起始列
offset4 = LOSE1col - schemaCol4   '赔1起始列-模式四起始列


'求模式五的计算公式

offset1 = dataCol(0) - schemaCol5    '威廉希尔起始列-模式五起始列
offset2 = dataCol(3) - schemaCol5    'Bet365起始列-模式五起始列
offset3 = dataCol(4) - schemaCol5    '澳门网起始列-模式五起始列
offset4 = LOSE1col - schemaCol5   '赔1起始列-模式五起始列


'求模式七的计算公式

offset1 = dataCol(0) - schemaCol7    '威廉希尔起始列-模式七起始列
offset2 = dataCol(1) - schemaCol7    'Bet365起始列-模式七起始列
offset3 = dataCol(2) - schemaCol7    '澳门网起始列-模式七起始列
offset4 = lose2Col - schemaCol7   '赔1起始列-模式七起始列


'求模式八的计算公式

offset1 = dataCol(0) - schemaCol8    '威廉希尔起始列-模式八起始列
offset2 = dataCol(3) - schemaCol8    'Bet365起始列-模式八起始列
offset3 = dataCol(4) - schemaCol8    '澳门网起始列-模式八起始列
offset4 = lose2Col - schemaCol8   '赔1起始列-模式八起始列

'数据回填至EXCEL中
For i = 0 To UBound(data)
    j = data(i, 0)     '取数据对应的worksheet中的行号
    '处理模式计算部分和bf1、bf2、bf3以及方差的特殊处理
    If data(i, 4) = "初始值" Then
        dataSheet.Cells(j, schemaCol1) = data(i, 26)        '模式一
        dataSheet.Cells(j, schemaCol2) = data(i, 27)    '模式二
        dataSheet.Cells(j, schemaCol3) = data(i, 28)    '模式三
        
        dataSheet.Cells(j + 1, schemaCol1) = data(i + 1, 26)
        dataSheet.Cells(j + 1, schemaCol2) = data(i + 1, 27)
        dataSheet.Cells(j + 1, schemaCol3) = data(i + 1, 28)
        
        dataSheet.Cells(j + 2, schemaCol1) = data(i + 2, 26)
        dataSheet.Cells(j + 2, schemaCol2) = data(i + 2, 27)
        dataSheet.Cells(j + 2, schemaCol3) = data(i + 2, 28)
        
        
        '模式6
        If Not (dataSheet.Cells(j, panmCol + 1) = "" And dataSheet.Cells(j + 1, panmCol + 1) = "" And dataSheet.Cells(j + 2, panmCol + 1) = "") Then
            dataSheet.Cells(j, schemaCol6) = dataSheet.Cells(j, panmCol + 1).Text + ":" + dataSheet.Cells(j + 1, panmCol + 1).Text + ":" + dataSheet.Cells(j + 2, panmCol + 1).Text
        End If
        dataSheet.Cells(j + 1, schemaCol6) = dataSheet.Cells(j, schemaCol6)
        dataSheet.Cells(j + 2, schemaCol6) = dataSheet.Cells(j, schemaCol6)
        
        

        
        
        '----------add by  ljqu 2018.8.20 begin------------------------
        '直接利用数据进行计算，不在通过公式来计算
        '威廉：5-7； Bet365：8-10； 澳门：11-13； 立博：14-16； 易胜博：17-19； 赔1:20-22；赔2:23-25；  模式一：26，模式2:27，模式3:28
        
        '模式四
        
        For k1 = 0 To 2       '初始值、即时一、即时二
            value1 = calDispersion(data(i + k1, 5), data(i + k1, 8), data(i + k1, 11), data(i + k1, 20)) '胜
            value2 = calDispersion(data(i + k1, 6), data(i + k1, 9), data(i + k1, 12), data(i + k1, 21))    '平
            value3 = calDispersion(data(i + k1, 7), data(i + k1, 10), data(i + k1, 13), data(i + k1, 22)) '负
            
            dataSheet.Cells(j + k1, schemaCol4) = 模式四(value1, value2, value3)   '模式四
            dataSheet.Cells(j + k1, schemaCol4 + 1) = ConcateData(value1, value2, value3, 4, 100) '模式四的值
        Next
        
        '模式五
        
        For k1 = 0 To 2       '初始值、即时一、即时二
            value1 = calDispersion(data(i + k1, 5), data(i + k1, 14), data(i + k1, 17), data(i + k1, 20))   '胜
            value2 = calDispersion(data(i + k1, 6), data(i + k1, 15), data(i + k1, 18), data(i + k1, 21))    '平
            value3 = calDispersion(data(i + k1, 7), data(i + k1, 16), data(i + k1, 19), data(i + k1, 22))    '负
            
            dataSheet.Cells(j + k1, schemaCol5) = 模式四(value1, value2, value3)   '模式四
            dataSheet.Cells(j + k1, schemaCol5 + 1) = ConcateData(value1, value2, value3, 4, 100) '模式四的值
        Next
        
        '模式七
        
        For k1 = 0 To 2       '初始值、即时一、即时二
            value1 = calDispersion(data(i + k1, 5), data(i + k1, 8), data(i + k1, 11), data(i + k1, 23)) '胜
            value2 = calDispersion(data(i + k1, 6), data(i + k1, 9), data(i + k1, 12), data(i + k1, 24))   '平
            value3 = calDispersion(data(i + k1, 7), data(i + k1, 10), data(i + k1, 13), data(i + k1, 25))    '负
            
            dataSheet.Cells(j + k1, schemaCol7) = 模式四(value1, value2, value3)   '模式四
            dataSheet.Cells(j + k1, schemaCol7 + 1) = ConcateData(value1, value2, value3, 4, 100) '模式四的值
        Next
        
        '模式八
        
        For k1 = 0 To 2       '初始值、即时一、即时二
            value1 = calDispersion(data(i + k1, 5), data(i + k1, 14), data(i + k1, 17), data(i + k1, 23))  '胜
            value2 = calDispersion(data(i + k1, 6), data(i + k1, 15), data(i + k1, 18), data(i + k1, 24))   '平
            value3 = calDispersion(data(i + k1, 7), data(i + k1, 16), data(i + k1, 19), data(i + k1, 25))    '负
            
            dataSheet.Cells(j + k1, schemaCol8) = 模式四(value1, value2, value3)   '模式四
            dataSheet.Cells(j + k1, schemaCol8 + 1) = ConcateData(value1, value2, value3, 4, 100) '模式四的值
        Next
        
        '---------add by ljqu 2018.8.20  end---------------------------
        
        
        
        'add by ljqu 2018.8.13 将模式7的数据并排为模式7并排
        dataSheet.Cells(j, schemaCol71) = dataSheet.Cells(j, schemaCol7)
        dataSheet.Cells(j + 1, schemaCol71) = dataSheet.Cells(j, schemaCol7)
        dataSheet.Cells(j + 2, schemaCol71) = dataSheet.Cells(j, schemaCol7)
        
        dataSheet.Cells(j, schemaCol71 + 1) = dataSheet.Cells(j + 1, schemaCol7)
        dataSheet.Cells(j + 1, schemaCol71 + 1) = dataSheet.Cells(j + 1, schemaCol7)
        dataSheet.Cells(j + 2, schemaCol71 + 1) = dataSheet.Cells(j + 1, schemaCol7)

        dataSheet.Cells(j, schemaCol71 + 2) = dataSheet.Cells(j + 2, schemaCol7)
        dataSheet.Cells(j + 1, schemaCol71 + 2) = dataSheet.Cells(j + 2, schemaCol7)
        dataSheet.Cells(j + 2, schemaCol71 + 2) = dataSheet.Cells(j + 2, schemaCol7)

        'add by ljqu 2018.8.13 将模式8的数据并排为模式8并排
        dataSheet.Cells(j, schemaCol81) = dataSheet.Cells(j, schemaCol8)
        dataSheet.Cells(j + 1, schemaCol81) = dataSheet.Cells(j, schemaCol8)
        dataSheet.Cells(j + 2, schemaCol81) = dataSheet.Cells(j, schemaCol8)
        
        dataSheet.Cells(j, schemaCol81 + 1) = dataSheet.Cells(j + 1, schemaCol8)
        dataSheet.Cells(j + 1, schemaCol81 + 1) = dataSheet.Cells(j + 1, schemaCol8)
        dataSheet.Cells(j + 2, schemaCol81 + 1) = dataSheet.Cells(j + 1, schemaCol8)

        dataSheet.Cells(j, schemaCol81 + 2) = dataSheet.Cells(j + 2, schemaCol8)
        dataSheet.Cells(j + 1, schemaCol81 + 2) = dataSheet.Cells(j + 2, schemaCol8)
        dataSheet.Cells(j + 2, schemaCol81 + 2) = dataSheet.Cells(j + 2, schemaCol8)

        
        
        '计算模式四至模式八比较值
        For k1 = 1 To 2
            dataSheet.Cells(j + k1, schemaCol4 + 2) = MethodCompare(dataSheet.Cells(j + k1, schemaCol4 + 1), dataSheet.Cells(j + k1 - 1, schemaCol4 + 1))
            dataSheet.Cells(j + k1, schemaCol5 + 2) = MethodCompare(dataSheet.Cells(j + k1, schemaCol5 + 1), dataSheet.Cells(j + k1 - 1, schemaCol5 + 1))
            dataSheet.Cells(j + k1, schemaCol7 + 2) = MethodCompare(dataSheet.Cells(j + k1, schemaCol7 + 1), dataSheet.Cells(j + k1 - 1, schemaCol7 + 1))
            dataSheet.Cells(j + k1, schemaCol8 + 2) = MethodCompare(dataSheet.Cells(j + k1, schemaCol8 + 1), dataSheet.Cells(j + k1 - 1, schemaCol8 + 1))
        Next

        '计算模式七中的“四七比较”和模式八中的“五八比较” add by ljqu 2018.3.18
        For k1 = 0 To 2
            dataSheet.Cells(j + k1, schemaCol7 + 3) = MethodCompare(dataSheet.Cells(j + k1, schemaCol7 + 1), dataSheet.Cells(j + k1, schemaCol4 + 1))
            dataSheet.Cells(j + k1, schemaCol8 + 3) = MethodCompare(dataSheet.Cells(j + k1, schemaCol8 + 1), dataSheet.Cells(j + k1, schemaCol5 + 1))
        Next
        
        '更正赔2的初始值的“比较”栏目的值
        If dataSheet.Cells(j, lose2Col) <> "" And dataSheet.Cells(j, lose2Col + 1) <> "" And dataSheet.Cells(j, lose2Col + 2) <> "" Then
            dataSheet.Cells(j, lose2Col + 4) = 横向比较(dataSheet, j, lose2Col + 3, lose2Col - LOSE1col, "A", 1)
        End If
        
        '计算bf1的初始值的"比较"栏数据
        
        If dataSheet.Cells(j, BF1Col) <> "" And dataSheet.Cells(j, BF1Col + 1) <> "" And dataSheet.Cells(j, BF1Col + 2) <> "" Then
            dataSheet.Cells(j, BF1Col + 5) = 横向比较(dataSheet, j, BF1Col + 5, 2, "D", 2)
        End If
        
        
        
        '重新计算方差的“标识”数据
        If dataSheet.Cells(j, varCol) <> "" And dataSheet.Cells(j, varCol + 1) <> "" And dataSheet.Cells(j, varCol + 2) <> "" Then
            dataSheet.Cells(j, varCol + 3) = 标识(dataSheet.Cells(j, varCol), dataSheet.Cells(j, varCol + 1), dataSheet.Cells(j, varCol + 2)) + 固定值比较(dataSheet.Cells(j, varCol), dataSheet.Cells(j, varCol + 1), dataSheet.Cells(j, varCol + 2), 1, "D")
        End If
        If dataSheet.Cells(j + 1, varCol) <> "" And dataSheet.Cells(j + 1, varCol + 1) <> "" And dataSheet.Cells(j + 1, varCol + 2) <> "" Then
            dataSheet.Cells(j + 1, varCol + 3) = 标识(dataSheet.Cells(j + 1, varCol), dataSheet.Cells(j + 1, varCol + 1), dataSheet.Cells(j + 1, varCol + 2)) + 固定值比较(dataSheet.Cells(j + 1, varCol), dataSheet.Cells(j + 1, varCol + 1), dataSheet.Cells(j + 1, varCol + 2), 1, "D")
        End If
        If dataSheet.Cells(j + 2, varCol) <> "" And dataSheet.Cells(j + 2, varCol + 1) <> "" And dataSheet.Cells(j + 2, varCol + 2) <> "" Then
            dataSheet.Cells(j + 2, varCol + 3) = 标识(dataSheet.Cells(j + 2, varCol), dataSheet.Cells(j + 2, varCol + 1), dataSheet.Cells(j + 2, varCol + 2)) + 固定值比较(dataSheet.Cells(j + 2, varCol), dataSheet.Cells(j + 2, varCol + 1), dataSheet.Cells(j + 2, varCol + 2), 1, "D")
        End If
        
        
        
        
        '-----------------------------------------
        'add by ljqu 2018.8.13
        '-----------------------------------------
        If Not (dataSheet.Cells(j, panbCol + 1) = "" And dataSheet.Cells(j + 1, panbCol + 1) = "" And dataSheet.Cells(j + 2, panbCol + 1) = "") Then
            dataSheet.Cells(j, schemaCol61) = dataSheet.Cells(j, panbCol + 1).Text + ":" + dataSheet.Cells(j + 1, panbCol + 1).Text + ":" + dataSheet.Cells(j + 2, panbCol + 1).Text
        End If
        dataSheet.Cells(j + 1, schemaCol61) = dataSheet.Cells(j, schemaCol61)
        dataSheet.Cells(j + 2, schemaCol61) = dataSheet.Cells(j, schemaCol61)
        
        
        'Ok30列转行
        Call 通用列转行(dataSheet, j, dataColDict, "OK30", 4, 1, 1, "OK30_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "OK30", 4, 1, 2, "OK30_1", 2)
        Call 通用列转行(dataSheet, j, dataColDict, "OK30", 4, 1, 3, "OK30_1", 3)
        Call 通用列转行(dataSheet, j, dataColDict, "OK30", 5, 1, 2, "OK30_2", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "OK30", 5, 1, 3, "OK30_2", 2)
        
        
        'Bf1列转行
        Call 通用列转行(dataSheet, j, dataColDict, "BF1", 5, 1, 1, "BF1_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "BF1", 5, 1, 2, "BF1_1", 2)
        Call 通用列转行(dataSheet, j, dataColDict, "BF1", 5, 1, 3, "BF1_1", 3)
        
        Call 通用列转行(dataSheet, j, dataColDict, "BF1", 6, 1, 1, "BF1_2", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "BF1", 6, 1, 2, "BF1_2", 2)
        Call 通用列转行(dataSheet, j, dataColDict, "BF1", 6, 1, 3, "BF1_2", 3)

        
        '威廉数据列转行
        Call 通用列转行(dataSheet, j, dataColDict, "DATAW", 1, 4, 1, "DATAW_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "DATAW", 1, 4, 2, "DATAW_2", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "DATAW", 1, 4, 3, "DATAW_3", 1)
        
        
        '-----------------------------------------
        'add end  2018.8.13
        '-----------------------------------------
        'Bet365数据列转行
        Call 通用列转行(dataSheet, j, dataColDict, "DATAB", 1, 4, 1, "DATAB_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "DATAB", 1, 4, 2, "DATAB_2", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "DATAB", 1, 4, 3, "DATAB_3", 1)
        
        
        '澳门彩票数据列转行
        Call 通用列转行(dataSheet, j, dataColDict, "DATAM", 1, 4, 1, "DATAM_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "DATAM", 1, 4, 2, "DATAM_2", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "DATAM", 1, 4, 3, "DATAM_3", 1)
        
        
        '赔1数据列转行
        Call 通用列转行(dataSheet, j, dataColDict, "LOSE1", 1, 3, 1, "LOSE1_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "LOSE1", 1, 3, 2, "LOSE1_2", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "LOSE1", 1, 3, 3, "LOSE1_3", 1)
        
        
        '赔2数据列转行
        Call 通用列转行(dataSheet, j, dataColDict, "LOSE2", 1, 3, 1, "LOSE2_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "LOSE2", 1, 3, 2, "LOSE2_2", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "LOSE2", 1, 3, 3, "LOSE2_3", 1)
        
        '----------------------------------------
        'add begin 2018.8.27
        '----------------------------------------
        
        '----------------------------------------
        'add end 2018.8.27
        '----------------------------------------


        '-----------------------------------------
        'add by ljqu 2018.8.18
        '-----------------------------------------
        
        
        '模式一列转行
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA1", 1, 1, 1, "SCHEMA1_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA1", 1, 1, 2, "SCHEMA1_1", 2)
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA1", 1, 1, 3, "SCHEMA1_1", 3)
        
        
        '模式二列转行
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA2", 1, 1, 1, "SCHEMA2_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA2", 1, 1, 2, "SCHEMA2_1", 2)
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA2", 1, 1, 3, "SCHEMA2_1", 3)
        
        
        '模式四列转行
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA4", 1, 1, 1, "SCHEMA4_1", 1)
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA4", 1, 1, 2, "SCHEMA4_1", 2)
        Call 通用列转行(dataSheet, j, dataColDict, "SCHEMA4", 1, 1, 3, "SCHEMA4_1", 3)
        
        
        '-----------------------------------------
        'add end  2018.8.18
        '-----------------------------------------

        
        
        
    End If
Next


MsgBox ("模式计算完毕！")

End Sub

Sub 模式计算一(data, cols As String, colDes As String, bgCol As Integer, methCol As Integer, schemaType As Integer)
'--------------------------------------------
' data:数据源
' cols:待比较的数据组号
' colDes:要生成的数据组标识
' bgCol:数据组开始的位置
' methCol:模式数据存放到数据源的位置
'schemaType:计算模式
'--------------------------------------------

Dim col
Dim colDesc
Dim colLen As Long
Dim k
Dim i, j
Dim cnt

Dim sortData()     '保存要排序的数据，第一行保存的是主胜数据，第二行保存的是平和数据，第三行是客胜数据
Dim sortIndex()    '保存排序后的索引

k = bgCol


col = Split(cols, ",")
colDesc = Split(colDes, ",")
colLen = UBound(col)
If colLen <> UBound(colDesc) And colLen <> 3 Then
    MsgBox ("模式一计算的对应参数不匹配！(列和说明不匹配）")
Else
    ReDim sortData(2, colLen)  '保存待计算的数据
    ReDim sortIndex(2, colLen)   '排序后的索引
    For i = 0 To UBound(data)
        '2017.10.15  由于即时值一和即时值二也需要按各计算各自的模式，所以此处的判断将改为只判断三个数值是否存在，
        'If data(i, 4) = "初始值" Then
            '移动数据到计算数组中
            For cnt = 0 To colLen
                j = CInt(col(cnt))
                sortData(0, cnt) = data(i, k + 3 * j)
                sortData(1, cnt) = data(i, k + 3 * j + 1)
                sortData(2, cnt) = data(i, k + 3 * j + 2)
            Next
            '对数据进行排序，返回排序后的数组序号
            Call SortSchemaData(sortData, sortIndex, "D")
            
            If schemaType = 2 Then
                data(i, methCol) = 生成模式符号(sortData, sortIndex, colDesc, 2)
            Else
                If sortData(0, 0) > sortData(2, 0) Then '判断第一组数据的主胜：主负
                    '判断三组数据的主胜
                    data(i, methCol) = 生成模式符号(sortData, sortIndex, colDesc, 1, 0)
                Else
                    '判断三组数据的主负
                    data(i, methCol) = "-" + 生成模式符号(sortData, sortIndex, colDesc, 1, 2)
                End If
            End If
        'End If
    Next
End If


End Sub
Function 生成模式符号(iDataSort, iSortIndex, colDesc, schemaType As Integer, Optional rowNo As Integer = 0)
'-----------------------------------------------------------------------
'iSortData：待排序的数据，用于比较数据是否相等
'iSortIndex：排序后的索引
'colDesc：对应模式说明字段
'SchemaType：模式类型。1：模式一和模式三表示方式，2：模式二表示方式，模式二中的colDesc保存的是"3,1,0"
'             3：模式四表示方式
'rowNo：生成数据的行号
'-----------------------------------------------------------------------
Dim rowLen, colLen
Dim indexNo
Dim schemaMsg As String
Dim i, j
Dim isOk As Boolean


rowLen = UBound(iSortIndex, 1)
colLen = UBound(iSortIndex, 2)
schemaMsg = ""
'模式一和模式三的模式描述
If schemaType = 1 Then

    For i = 0 To colLen - 1
        indexNo = iSortIndex(rowNo, i)  '获取第一个序列号
        schemaMsg = schemaMsg + colDesc(indexNo)
        If iDataSort(rowNo, indexNo) = iDataSort(rowNo, iSortIndex(rowNo, i + 1)) Then   '如果前后两个数据相等，则加上“=”号
            schemaMsg = schemaMsg + "="
        End If
    Next
    '取最后一个
    indexNo = iSortIndex(rowNo, i)
    schemaMsg = schemaMsg + colDesc(indexNo)
End If

'模式二的模式描述
If schemaType = 2 Then
    For i = 0 To colLen    '行代表的主胜、平、客胜
        isOk = False
        For j = 0 To rowLen
            If iSortIndex(j, 0) = i Then
                schemaMsg = schemaMsg + colDesc(j)
                isOk = True
            End If
        Next
        If Not isOk Then    '中间没有该组数据
            schemaMsg = schemaMsg + "X"
        End If
        If i < colLen Then      '最后一个数据后面不用加分隔符
                schemaMsg = schemaMsg + ":"
        End If
    Next
End If


'比较的模式描述
If schemaType = 4 Then
    For i = 0 To colLen
        indexNo = iSortIndex(rowNo, i)  '获取第一个序列号
        If iDataSort(rowNo, indexNo) <> 0 Then schemaMsg = schemaMsg + colDesc(indexNo)
    Next
End If

生成模式符号 = schemaMsg
End Function

Sub SortSchemaData(iSortData, sortIndex, Optional sortType As String = "A", Optional RowOrCol As String = "R")
'对几组数据进行排序，
'sortData 是待排序的二维数组
'sortIndex 是保存排序后的索引
'rowOrCol: 行列排序说明：R:对行进行排序，C：对列进行排序
'sortType 是排序类型：A：升序，D：降序
Dim i, j, k
Dim rowLen, colLen
Dim tempData
Dim tempIndex
Dim temp
Dim sortData1()

sortData1 = iSortData

rowLen = UBound(sortData1, 1)
colLen = UBound(sortData1, 2)


If RowOrCol = "R" Then
    For i = 0 To rowLen
         '以下对每一行的数据进行排序，排序结果将数据序号保存在sortIndex对应的行中
        '初始化索引数组
        For j = 0 To colLen
            sortIndex(i, j) = j
        Next
        For j = 0 To colLen
            tempData = sortData1(i, j)
            tempIndex = j
            For k = j + 1 To colLen
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
            
            '交换数据
            If tempIndex <> j Then
                sortData1(i, tempIndex) = sortData1(i, j)
                sortData1(i, j) = tempData
                temp = sortIndex(i, tempIndex)
                sortIndex(i, tempIndex) = sortIndex(i, j)
                sortIndex(i, j) = temp
                'sortIndex(i, j) = tempIndex
                'sortData1(i, tempIndex) = 0
            End If
        Next
    Next
ElseIf RowOrCol = "C" Then
    For i = 0 To colLen
         '以下对每一行的数据进行排序，排序结果将数据序号保存在sortIndex对应的行中
        '初始化索引数组
        For j = 0 To rowLen
            sortIndex(j, i) = j
        Next
        For j = 0 To rowLen
            tempData = sortData1(j, i)
            tempIndex = j
            For k = j + 1 To rowLen
                If sortType = "D" Then   '降序
                    If sortData1(k, i) > tempData Then
                        tempData = sortData1(k, i)
                        tempIndex = k
                    End If
                Else       '默认升序
                    If sortData1(k, i) < tempData Then
                        tempData = sortData1(k, i)
                        tempIndex = k
                    End If

                End If
            Next
            
            '交换数据
            If tempIndex <> j Then
                sortData1(tempIndex, i) = sortData1(j, i)
                sortData1(j, i) = tempData
                temp = sortIndex(tempIndex, i)
                sortIndex(tempIndex, i) = sortIndex(j, i)
                sortIndex(j, i) = temp
            End If
        Next
    Next
End If

End Sub



Sub SortSchemaData_Old(iSortData, sortIndex, Optional sortType As String = "A", Optional RowOrCol As String = "R")
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


If RowOrCol = "R" Then
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
            sortData1(i, tempIndex) = 0
        Next
    Next
ElseIf RowOrCol = "C" Then
    For i = 0 To colLen
         '以下对每一行的数据进行排序，排序结果将数据序号保存在sortIndex对应的行中
        For j = 0 To rowLen
            tempData = sortData1(j, i)
            tempIndex = j
            For k = 0 To rowLen
            
                If sortType = "D" Then   '降序
                    If sortData1(k, i) > tempData Then
                        tempData = sortData1(k, i)
                        tempIndex = k
                    End If
                Else       '默认升序
                    If sortData1(k, i) < tempData Then
                        tempData = sortData1(k, i)
                        tempIndex = k
                    End If

                End If
            Next
            sortIndex(j, i) = tempIndex
            sortData1(tempIndex, i) = 0
        Next
    Next
End If

End Sub


Function 赛季计算(dt)
Dim yyyy As Integer
Dim mm As Integer

yyyy = Year(dt) Mod 100
mm = Month(dt)

If mm <= 5 Then
    赛季计算 = Trim(str(yyyy - 1)) + "-" + Trim(str(yyyy)) + "赛季"
ElseIf mm >= 8 Then
    赛季计算 = Trim(str(yyyy)) + "-" + Trim(str(yyyy + 1)) + "赛季"
End If


End Function

Sub 球探网初始()
    数据初始
End Sub


Sub 球探网更新()
    数据更新
End Sub

Sub 记录初始值(dataSheet As Worksheet, dataW, Loc, dataWcol, i, bgCol As Integer, realBgCol As Integer, Optional isRealRec As Boolean = True, Optional isLabel As Boolean = False, Optional compareType As String = "A", Optional cnt As Integer = 3)
'------------------------------------------------------------
'跟保存SHEET相关的输入变量
'  dataSheet 待记录的工作表
'  loc  工作表的位置
'  dataWCol 数据存放开始列
'跟操作数组相关的变量
'  dataW 数据数组
'  i   数据数组的行指针
'  bgCol： 数据数组中要何存的初始数据的起始列号,     球探网的数据是从第7列开始的
'  realBgCol:数据数组中要存放的即时候数据的起始列号，     add by ljqu 2015.3.29,   增加这一列便于处理初始数据和实时数据的存放不是固定的排列顺序，原来默认即时数据跟在实始数据之后
'            变更后，将原来默认的即时值坐标由bgCol+cnt改为realBgCol
'
'  isRealRec：是否记录即时值，true：记录，false：不记录
'  isLabel：是否记录标识值，True:记录，false:不记录，对于赔率数据只有比较数据，默认为false，对于凯利数据则设为true
'  CompareType：是否记录比较值，"A":按升序比较，"D":按降序，"FALSE"：没有比较栏
'  cnt：要过录的数据列数，默认3列，BF1、BF2、BF3是四列
'------------------------------------------------------------
Dim upColor
Dim downColor
Dim j, j1
Dim formulaStr As String

upColor = vbRed
downColor = vbGreen

         '--初始值
         For j1 = 0 To cnt - 1
            dataSheet.Cells(Loc, dataWcol + j1) = dataW(i, bgCol + j1)
            dataSheet.Cells(Loc, dataWcol + j1).Font.Color = vbBlack
        Next
        j = cnt
        If isLabel Then
            'dataSheet.Cells(loc, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
            dataSheet.Cells(Loc, dataWcol + j) = 标识(dataSheet.Cells(Loc, dataWcol), dataSheet.Cells(Loc, dataWcol + 1), dataSheet.Cells(Loc, dataWcol + 2))
            j = j + 1
        End If
        
        
        If isRealRec Then
            ' --即时值1
            For j1 = 0 To cnt - 1
                If dataW(i, realBgCol + j1) = "" Then
                    dataSheet.Cells(Loc + 1, dataWcol + j1) = dataSheet.Cells(Loc, dataWcol + j1)
                Else
                    dataSheet.Cells(Loc + 1, dataWcol + j1) = dataW(i, realBgCol + j1)
                End If
                 '颜色区分
                If dataSheet.Cells(Loc, dataWcol + j1) > dataSheet.Cells(Loc + 1, dataWcol + j1) Then
                    dataSheet.Cells(Loc + 1, dataWcol + j1).Font.Color = downColor
                ElseIf dataSheet.Cells(Loc, dataWcol + j1) < dataSheet.Cells(Loc + 1, dataWcol + j1) Then
                    dataSheet.Cells(Loc + 1, dataWcol + j1).Font.Color = upColor
                Else
                    dataSheet.Cells(Loc + 1, dataWcol + j1).Font.Color = vbBlack
                End If
            Next

            j = cnt
            If isLabel Then
                'dataSheet.Cells(loc + 1, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
                dataSheet.Cells(Loc + 1, dataWcol + j) = 标识(dataSheet.Cells(Loc + 1, dataWcol), dataSheet.Cells(Loc + 1, dataWcol + 1), dataSheet.Cells(Loc + 1, dataWcol + 2))
                j = j + 1
            End If
            '计算比较列
            
            '计算比较列
            If compareType = "A" Or compareType = "D" Then
                Call 比较(dataSheet, Loc + 1, dataWcol + j, j, compareType)
            End If

            '即时值二
            For j1 = 0 To cnt - 1
                If dataW(i, realBgCol + j1) = "" Then
                    dataSheet.Cells(Loc + 2, dataWcol + j1) = dataSheet.Cells(Loc, dataWcol + j1)
                Else
                    dataSheet.Cells(Loc + 2, dataWcol + j1) = dataW(i, realBgCol + j1)
                End If
                
            Next
            
            j = cnt
            If isLabel Then
                'dataSheet.Cells(loc + 2, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
                dataSheet.Cells(Loc + 2, dataWcol + j) = 标识(dataSheet.Cells(Loc + 2, dataWcol), dataSheet.Cells(Loc + 2, dataWcol + 1), dataSheet.Cells(Loc + 2, dataWcol + 2))
                j = j + 1
            End If
            '计算比较列
            If compareType = "A" Or compareType = "D" Then
                Call 比较(dataSheet, Loc + 2, dataWcol + j, j, compareType)
            End If
        End If
End Sub


Sub 记录即时值(dataSheet As Worksheet, dataW, Loc, dataWcol, i, initialBgCol As Integer, bgCol As Integer, Optional isLabel As Boolean = False, Optional compareType As String = "A", Optional cnt As Integer = 3)
'------------------------------------------------------------
'  dataSheet 待记录的工作表
'  dataW 数据数组
'  loc  工作表的位置
'  dataWCol 数据存放开始列
'  i   数据数组的指针
'  initialBgCol:数据数组中要存放的即时候数据的起始列号，     add by ljqu 2015.3.29,   增加这一列便于处理初始数据和实时数据的存放不是固定的排列顺序，原来默认即时数据跟在实始数据之后
'            变更后，将原来默认的即时值坐标由bgCol -cnt 改为initialBgCol
'
'  bgCol： 数据数组中要何存的数据的起始列号,     球探网的数据是从第10列开始的

'  isLabel：是否记录标识值，True:记录，false:不记录，对于赔率数据只有比较数据，默认为false，对于凯利数据则设为true
'  CompareType：是否记录比较值，"A":按升序比较，"D":按降序
'  cnt：要过录的数据列数，默认3列，BF1、BF2、BF3是四列
'------------------------------------------------------------
Dim upColor
Dim downColor
Dim j, j1
Dim formulaStr As String

Dim n1, n2, n3
'upColor = Dict.Item("REAL_UPCOLOR")
'downColor = Dict.Item("REAL_DOWNCOLOR")
upColor = vbRed
downColor = vbGreen
    If dataW(i, bgCol) <> "" And dataW(i, bgCol + 1) <> "" And dataW(i, bgCol + 2) <> "" Then
        '如果初始值为空，也同时更新初始值
        If dataSheet.Cells(Loc - 2, dataWcol) = "" Or dataSheet.Cells(Loc - 2, dataWcol + 1) = "" Or dataSheet.Cells(Loc - 2, dataWcol + 2) = "" Then
        
            '过录初始值
            For j1 = 0 To cnt - 1
                dataSheet.Cells(Loc - 2, dataWcol + j1) = dataW(i, initialBgCol + j1)
                dataSheet.Cells(Loc - 2, dataWcol + j1).Font.Color = vbBlack
            Next
            j = cnt
            If isLabel Then
                'dataSheet.Cells(loc - 2, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
                dataSheet.Cells(Loc - 2, dataWcol + j) = 标识(dataSheet.Cells(Loc - 2, dataWcol), dataSheet.Cells(Loc - 2, dataWcol + 1), dataSheet.Cells(Loc - 2, dataWcol + 2))
                j = j + 1
            End If

            
        End If
  
  
  
  
         '--即时值1为空，则将即时值填入即时值1和即时值2
  
        If dataSheet.Cells(Loc - 1, dataWcol) = "" Or dataSheet.Cells(Loc - 1, dataWcol + 1) = "" Or dataSheet.Cells(Loc - 1, dataWcol + 2) = "" Then
        
        
            '过录即时值一
            For j1 = 0 To cnt - 1
                dataSheet.Cells(Loc - 1, dataWcol + j1) = dataW(i, bgCol + j1)
                
                 '颜色区分
                If dataSheet.Cells(Loc - 2, dataWcol + j1) > dataSheet.Cells(Loc - 1, dataWcol + j1) Then
                    dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = downColor
                ElseIf dataSheet.Cells(Loc - 2, dataWcol + j1) < dataSheet.Cells(Loc - 1, dataWcol + j1) Then
                    dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = upColor
                Else
                    dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = vbBlack
                End If
                
            Next
            
            j = cnt
            If isLabel Then
                'dataSheet.Cells(loc - 1, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
                dataSheet.Cells(Loc - 1, dataWcol + j) = 标识(dataSheet.Cells(Loc - 1, dataWcol), dataSheet.Cells(Loc - 1, dataWcol + 1), dataSheet.Cells(Loc - 1, dataWcol + 2))
                j = j + 1
            End If
            
            '计算比较列
            If compareType = "A" Or compareType = "D" Then
                Call 比较(dataSheet, Loc - 1, dataWcol + j, j, compareType)
            End If
            
             '过录即时值二
            For j1 = 0 To cnt - 1
                dataSheet.Cells(Loc, dataWcol + j1) = dataW(i, bgCol + j1)
                dataSheet.Cells(Loc, dataWcol + j1).Font.Color = vbBlack
            Next
            
            j = cnt
            If isLabel Then
                dataSheet.Cells(Loc, dataWcol + j) = 标识(dataSheet.Cells(Loc, dataWcol), dataSheet.Cells(Loc, dataWcol + 1), dataSheet.Cells(Loc, dataWcol + 2))
                j = j + 1
            End If

            
        Else
            
            For j1 = 0 To cnt - 1
            
                If dataSheet.Cells(Loc, dataWcol + j1).Value <> dataW(i, bgCol + j1) Then
                    dataSheet.Cells(Loc - 1, dataWcol + j1) = dataSheet.Cells(Loc, dataWcol + j1)
                    dataSheet.Cells(Loc, dataWcol + j1) = dataW(i, bgCol + j1)
                   
                End If
                
                
                 '颜色区分---即时值一：初始值
                If dataSheet.Cells(Loc - 2, dataWcol + j1) > dataSheet.Cells(Loc - 1, dataWcol + j1) Then
                    dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = downColor
                ElseIf dataSheet.Cells(Loc - 2, dataWcol + j1) < dataSheet.Cells(Loc - 1, dataWcol + j1) Then
                    dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = upColor
                Else
                    dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = vbBlack
                End If
                
                 '颜色区分——即时值二：即时值
                If dataSheet.Cells(Loc - 1, dataWcol + j1) > dataSheet.Cells(Loc, dataWcol + j1) Then
                    dataSheet.Cells(Loc, dataWcol + j1).Font.Color = downColor
                ElseIf dataSheet.Cells(Loc - 1, dataWcol + j1) < dataSheet.Cells(Loc, dataWcol + j1) Then
                    dataSheet.Cells(Loc, dataWcol + j1).Font.Color = upColor
                Else
                    dataSheet.Cells(Loc, dataWcol + j1).Font.Color = vbBlack
                End If
            Next
            
            '计算标识和比较列
            '------即时值一
            j = cnt
            If isLabel Then
                'dataSheet.Cells(loc - 1, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
                dataSheet.Cells(Loc - 1, dataWcol + j) = 标识(dataSheet.Cells(Loc - 1, dataWcol), dataSheet.Cells(Loc - 1, dataWcol + 1), dataSheet.Cells(Loc - 1, dataWcol + 2))
                j = j + 1
            End If
            
            '计算比较列
            If compareType = "A" Or compareType = "D" Then
                Call 比较(dataSheet, Loc - 1, dataWcol + j, j, compareType)
            End If
            

            '------即时值二
            j = cnt
            If isLabel Then
                'dataSheet.Cells(loc, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
                dataSheet.Cells(Loc, dataWcol + j) = 标识(dataSheet.Cells(Loc, dataWcol), dataSheet.Cells(Loc, dataWcol + 1), dataSheet.Cells(Loc, dataWcol + 2))
                j = j + 1
            End If
            '计算比较列
            If compareType = "A" Or compareType = "D" Then
                Call 比较(dataSheet, Loc, dataWcol + j, j, compareType)
            End If
            
        End If
        
    End If

End Sub



Sub 处理竞彩网数据()
'此方法直接从网站截取数据
Dim x1 As Worksheet
Dim dataSheet As Worksheet
Dim cnt As Long
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim c1 As Integer   '用于记录让球的位置:列号

Dim str1 As String
Dim str2 As String
Dim str3
Dim str4
Dim currdate
Dim data()
Dim jcdata()

'定义赛事信息
Dim DataJcol    '中国竞彩网所在列
Dim gameDate    '赛事日期
Dim gameType    '赛事——联赛名称
Dim priTeam     '主队
Dim secTeam     '客队
Dim posiRatio   '主胜
Dim equalRatio  '平
Dim negRatio    '主负
Dim Loc         '定义主数据中的位置
Dim k1          '查找指针
Dim isRun       '是否正式运行状态，true:正式运行，false:调试状态

        '加载数据综合表
        Call 初始化字典(Dict, "Param")
        DataJcol = Dict.Item("DATAJ_COL")
        isRun = CBool(Dict.Item("IS_RUN"))
        
        Call 综合数据载入内存(data, "综合数据", , 3)
        
        '从网站加载竞彩网数据
        Call 加载中国竞彩网数据(jcdata, "中国竞彩网")
        
        Set dataSheet = ActiveWorkbook.Sheets("综合数据")
        

        
        Set x1 = ActiveWorkbook.Sheets("中国竞彩网")

        'cnt = x1.UsedRange.Rows(x1.UsedRange.Rows.Count).Row
        
        If isRun Then
            x1.Visible = xlSheetVeryHidden
        Else
            x1.Visible = xlSheetVisible
            x1.Select
            Selection.ClearContents
        End If
        
        x1.Cells(1, 1) = "ID"
        x1.Cells(1, 2) = "日期"
        x1.Cells(1, 3) = "编号"
        x1.Cells(1, 4) = "赛事"
        x1.Cells(1, 5) = "主队"
        x1.Cells(1, 6) = "客队"
        x1.Cells(1, 7) = "主胜"
        x1.Cells(1, 8) = "平"
        x1.Cells(1, 9) = "主负"
        
        For i = 1 To UBound(jcdata)
                '记录数据
                For j = 0 To UBound(jcdata, 2)
                    x1.Cells(i + 1, j + 1) = jcdata(i, j)
                Next
                '处理数据
                '在综合记录表中查找数据，找到则记录数据
              gameDate = CDate(jcdata(i, 1))
              gameType = jcdata(i, 3)
              priTeam = jcdata(i, 4)
              secTeam = jcdata(i, 5)
              posiRatio = jcdata(i, 6)
              equalRatio = jcdata(i, 7)
              negRatio = jcdata(i, 8)
              
              
              For k1 = 1 To UBound(data)
                  If gameDate = data(k1, 1) And gameType = data(k1, 7) And (priTeam = data(k1, 3) Or secTeam = data(k1, 4)) Then '日期，联赛，主队或客队
                      Exit For
                  End If
              Next
              
              If k1 <= UBound(data) Then
                  Loc = data(k1, 0)  '获取数据在sheet中的位置
                  dataSheet.Cells(Loc, DataJcol) = posiRatio
                  dataSheet.Cells(Loc, DataJcol + 1) = equalRatio
                  dataSheet.Cells(Loc, DataJcol + 2) = negRatio
                  dataSheet.Cells(Loc, DataJcol + 3).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
              End If
        
        Next
        MsgBox ("竞彩网数据处理完毕！")
End Sub


Sub 处理当期数据()
'
'
Dim row As Long
Dim col As Integer
Dim cc As Integer
Dim wkSheet As Worksheet
Dim dataSheet As Worksheet
Dim configData()
Dim i, j, k
Dim Loc As Integer
Dim dataBgCol As Integer   '指标数据开始列
Dim paraCnt As Integer     '基础数据个数
Dim OkNet As Integer       '澳客网当期编号

'加载配置数据
Call 配置数据载入(configData, "Config")

Call 初始化字典(Dict, "Param")

Set dataSheet = ActiveWorkbook.Sheets("综合数据")
Set wkSheet = ActiveWorkbook.Sheets("当期数据")

Call 初始化一般字典(dataColDict, dataSheet, 4, 0, 1, False)
OkNet = CInt(dataColDict.Item("OKID"))

wkSheet.Activate
wkSheet.Cells(1, 4) = CStr(dataSheet.Cells(1, 9))
wkSheet.Cells(2, 1) = "期数"
wkSheet.Cells(2, 2) = "编号"
wkSheet.Cells(2, 3) = "对阵"
wkSheet.Cells(2, 4) = "数据类型"
wkSheet.Cells(2, 5) = "联赛"
wkSheet.Cells(2, 6) = "模式1"
wkSheet.Cells(2, 7) = "模式2"
wkSheet.Cells(2, 8) = "模式3"
wkSheet.Cells(2, 9) = "模式4"
wkSheet.Cells(2, 10) = "模式5"
wkSheet.Cells(2, 11) = "模式6"
col = 12
For i = 2 To UBound(configData)
    If CBool(configData(i, 7)) Then
        wkSheet.Cells(2, col) = configData(i, 1) & Chr(10) & "标识"
        col = col + 1
        If configData(i, 8) = "A" Or configData(i, 8) = "D" Then
            wkSheet.Cells(2, col) = configData(i, 1) & Chr(10) & "比较"
            col = col + 1
        End If
    ElseIf configData(i, 8) = "A" Or configData(i, 8) = "D" Then
        wkSheet.Cells(2, col) = configData(i, 1)
        col = col + 1
    End If
Next i

row = dataSheet.UsedRange.Rows(dataSheet.UsedRange.Rows.Count).row

Loc = 3
For i = 5 To row
    If dataSheet.Cells(i, 1).EntireRow.Hidden = False Then
    '移植数据
        wkSheet.Cells(1, 3) = "日期：" & dataSheet.Cells(i, 1)
        '固定项移植
        wkSheet.Cells(Loc, 1) = dataSheet.Cells(i, OkNet)  '期数
        wkSheet.Cells(Loc, 2) = dataSheet.Cells(i, OkNet + 1) '编号
        wkSheet.Cells(Loc, 3) = dataSheet.Cells(i, 5)    '"对阵"
        wkSheet.Cells(Loc, 4) = dataSheet.Cells(i, 6)     '"数据类型"
        wkSheet.Cells(Loc, 5) = dataSheet.Cells(i, 7)     '"联赛"
        wkSheet.Cells(Loc, 6) = dataSheet.Cells(i, 13)     '"模式1"
        wkSheet.Cells(Loc, 7) = dataSheet.Cells(i, 14)     '"模式2"
        wkSheet.Cells(Loc, 8) = dataSheet.Cells(i, 15)     '"模式3"
        wkSheet.Cells(Loc, 9) = dataSheet.Cells(i, 16)     '"模式4"
        wkSheet.Cells(Loc, 10) = dataSheet.Cells(i, 17)     '"模式5"
        wkSheet.Cells(Loc, 11) = dataSheet.Cells(i, 18)    '模式6
        '指标值移植
        col = 12
        For j = 2 To UBound(configData)
            dataBgCol = CInt(dataColDict.Item(configData(j, 2)))
            paraCnt = CInt(configData(j, 9))
            If CBool(configData(j, 7)) Then
                wkSheet.Cells(Loc, col) = dataSheet.Cells(i, dataBgCol + paraCnt)
                col = col + 1
                If configData(j, 8) = "A" Or configData(j, 8) = "D" Then
                    wkSheet.Cells(Loc, col) = dataSheet.Cells(i, dataBgCol + paraCnt + 1)
                    col = col + 1
                End If
            ElseIf configData(j, 8) = "A" Or configData(j, 8) = "D" Then
                wkSheet.Cells(Loc, col) = dataSheet.Cells(i, dataBgCol + paraCnt)
                col = col + 1
            End If
        Next j
        Loc = Loc + 1
    End If
    
Next i

wkSheet.Cells.EntireColumn.AutoFit  'Columns("1:38").EntireColumn.AutoFit

'数据排序

    wkSheet.Range(Cells(2, 1), Cells(Loc - 1, col - 1)).Select
    Selection.Sort Key1:=Range(Cells(2, 1), Cells(Loc - 1, 1)), order1:=xlAscending, Header:=xlYes

    Set wkSheet = Nothing
    Set dataSheet = Nothing
End Sub






Sub 手工数据刷新(ByRef control As Office.IRibbonControl)
Dim data()
Dim configData()
Dim dataSheet As Worksheet
Dim cnt
Dim i, j, k

'跟参数配置相关的变量

Dim k1, j1        '参数配置处理中用至循环变量
Dim dataBeginCol As Integer     '数据开始的列号
Dim paraCnt As Integer       '基础数据长度
Dim paraCompType As String             '比较处理类型
Dim upColor
Dim downColor



Dim Loc As Long


Dim dataWcol As Integer   '威廉希尔数据开始列号
Dim dataBcol As Integer     'Bet365数据开始列号
Dim dataMcol As Integer     '澳门彩票数据开始列号
Dim dataLcol As Integer     '立博(英国)数据开始列号
Dim dataEcol As Integer     '易胜博数据开始列号
Dim LOSE1col As Integer      '赔1数据列号
Dim lose2Col As Integer      '赔2数据列号
Dim BF1Col As Integer        'BF1数据开始列号
Dim BF2col As Integer        'BF2数据开始列号
Dim BF3col As Integer        'BF3数据开始列号
Dim varCol As Integer        '方差数据开始列号

Dim recCnt  As Long     '待取的数据个数

Dim offset1 As Integer         '至第1个数据的偏移量
Dim offset2 As Integer         '至第2个数据的偏移量
Dim offset3 As Integer         '至第3个数据的偏移量
Dim offset4 As Integer         '至第4个数据的偏移量


Dim schemaCol As Integer
Dim schemaCol4 As Integer
Dim schemaCol5 As Integer

Dim dataCol(5)
Dim formulaStr As String    '模式四的计算公式
Dim formulaStr5 As String    '模式五的计算公式


upColor = vbRed
downColor = vbGreen

Call 初始化字典(Dict, "Param")

recCnt = CLng(Dict.Item("DEAL_RECCOUNT"))


Call 配置数据载入(configData, "Config")

Set dataSheet = ActiveWorkbook.Sheets("综合数据")
cnt = dataSheet.UsedRange.Rows(dataSheet.UsedRange.Rows.Count).row
ReDim data(cnt - 4, 23) '对应worksheet中的行号，日期，联赛，对阵，数据类型，五组数据（每组3个)，模式四组（1-4）

Call 初始化一般字典(dataColDict, dataSheet, 4, 0, 1, False)

dataCol(0) = dataColDict.Item("DATAW")
dataCol(1) = dataColDict.Item("DATAB")
dataCol(2) = dataColDict.Item("DATAM")
dataCol(3) = dataColDict.Item("DATAL")
dataCol(4) = dataColDict.Item("DATAE")
LOSE1col = dataColDict.Item("LOSE1")
lose2Col = dataColDict.Item("LOSE2")
BF1Col = dataColDict.Item("BF1")
BF2col = dataColDict.Item("BF2")
BF3col = dataColDict.Item("BF3")
varCol = dataColDict.Item("VAR")
schemaCol = dataColDict.Item("SCHEMA")

'将数据取到内存数组中

'确定取数的范围，为0则取全部
If recCnt = 0 Then
    i = 4
ElseIf cnt - recCnt + 1 < 4 Then
    i = 4
Else
    i = cnt - recCnt + 1
    ReDim data(recCnt - 1, 23) '对应worksheet中的行号，日期，联赛，对阵，数据类型，五组数据（每组3个)，模式四组（1-4）
End If

Loc = 0
Do While i <= cnt
    If dataSheet.Cells(i, 1) <> "" Then   'And dataSheet.Cells(i, 3) <> "" And dataSheet.Cells(i, 5) <> "" Then
        'If dataSheet.Cells(i, 6) = "初始值" Then
            data(Loc, 0) = i      '对应workSheet中的行号，用于在回填时直接写到对应的单元格
            data(Loc, 1) = dataSheet.Cells(i, 1) '日期
            data(Loc, 2) = dataSheet.Cells(i, 7) '联赛
            data(Loc, 3) = dataSheet.Cells(i, 5)  '对阵
            data(Loc, 4) = dataSheet.Cells(i, 6)  '数据类型
            k = 5
            For j = 0 To 4
                data(Loc, k + 3 * j) = dataSheet.Cells(i, dataCol(j))
                data(Loc, k + 3 * j + 1) = dataSheet.Cells(i, dataCol(j) + 1)
                data(Loc, k + 3 * j + 2) = dataSheet.Cells(i, dataCol(j) + 2)
            Next
            '指针移位
            Loc = Loc + 1
        '    i = i + 3
        'Else
            i = i + 1
        'End If
    Else
        Exit Do
    End If
Loop

Call 模式计算一(data, "0,1,2", "W,B,M", 5, 20, 1)  '模式一
Call 模式计算一(data, "0,1,2", "3,1,0", 5, 21, 2)  '模式二
Call 模式计算一(data, "0,4,3", "W,E,L", 5, 22, 1)   '模式三


'求模式四的计算公式
schemaCol4 = schemaCol + 3   '模式四的起始列
schemaCol5 = schemaCol + 4   '模式五的起始列

offset1 = dataCol(0) - schemaCol4    '威廉希尔起始列-模式四起始列
offset2 = dataCol(1) - schemaCol4    'Bet365起始列-模式四起始列
offset3 = dataCol(2) - schemaCol4    '澳门网起始列-模式四起始列
offset4 = LOSE1col - schemaCol4   '赔1起始列-模式四起始列

'生成模式四的脚本
formulaStr = "=模式四("
For i = 0 To 2
    formulaStr = formulaStr + "sum(RC[" + Trim(str(offset1 + i)) + "],RC[" + Trim(str(offset2 + i)) + "],RC[" + Trim(str(offset3 + i)) + "])/3-R[2]C[" + Trim(str(offset4 + i)) + "]" + IIf(i <> 2, ",", "")
Next
formulaStr = formulaStr + ")"


'求模式五的计算公式
schemaCol4 = schemaCol + 4   '模式五的起始列

offset1 = dataCol(0) - schemaCol4    '威廉希尔起始列-模式四起始列
offset2 = dataCol(3) - schemaCol4    'Bet365起始列-模式四起始列
offset3 = dataCol(4) - schemaCol4    '澳门网起始列-模式四起始列
offset4 = LOSE1col - schemaCol4   '赔1起始列-模式四起始列

'生成模式五的脚本
formulaStr5 = "=模式四("
For i = 0 To 2
    formulaStr5 = formulaStr5 + "sum(RC[" + Trim(str(offset1 + i)) + "],RC[" + Trim(str(offset2 + i)) + "],RC[" + Trim(str(offset3 + i)) + "])/3-R[2]C[" + Trim(str(offset4 + i)) + "]" + IIf(i <> 2, ",", "")
Next
formulaStr5 = formulaStr5 + ")"

'数据回填至EXCEL中
For i = 0 To UBound(data)
    j = data(i, 0)     '取数据对应的worksheet中的行号
    '处理模式计算部分和bf1、bf2、bf3以及方差的特殊处理
    If data(i, 4) = "初始值" Then
        dataSheet.Cells(j, schemaCol) = data(i, 20)
        dataSheet.Cells(j, schemaCol + 1) = data(i, 21)
        dataSheet.Cells(j, schemaCol + 2) = data(i, 22)
        dataSheet.Cells(j, schemaCol + 3).FormulaR1C1 = formulaStr    'data(i, 23)
        dataSheet.Cells(j, schemaCol + 4).FormulaR1C1 = formulaStr5    'data(i, 23)
        
        
        '同时复制到即时值1和即时值2
        dataSheet.Cells(j + 1, schemaCol) = data(i, 20)
        dataSheet.Cells(j + 1, schemaCol + 1) = data(i, 21)
        dataSheet.Cells(j + 1, schemaCol + 2) = data(i, 22)
        dataSheet.Cells(j + 1, schemaCol + 3) = dataSheet.Cells(j, schemaCol + 3)
        dataSheet.Cells(j + 1, schemaCol + 4) = dataSheet.Cells(j, schemaCol + 4)
        
        dataSheet.Cells(j + 2, schemaCol) = data(i, 20)
        dataSheet.Cells(j + 2, schemaCol + 1) = data(i, 21)
        dataSheet.Cells(j + 2, schemaCol + 2) = data(i, 22)
        dataSheet.Cells(j + 2, schemaCol + 3) = dataSheet.Cells(j, schemaCol + 3)
        dataSheet.Cells(j + 2, schemaCol + 4) = dataSheet.Cells(j, schemaCol + 4)
        
        dataSheet.Cells(j, schemaCol + 3) = dataSheet.Cells(j + 1, schemaCol + 3)
        dataSheet.Cells(j, schemaCol + 4) = dataSheet.Cells(j + 1, schemaCol + 4)
        
        
        
        '处理初始值中的标识列
        For k1 = 2 To UBound(configData)
            If CBool(configData(k1, 7)) Then          '是否具有标识列配置项
                paraCnt = CInt(configData(k1, 9))
                dataBeginCol = CInt(dataColDict.Item(configData(k1, 2)))
                dataSheet.Cells(j, dataBeginCol + paraCnt) = 标识(dataSheet.Cells(j, dataBeginCol), dataSheet.Cells(j, dataBeginCol + 1), dataSheet.Cells(j, dataBeginCol + 2))
            End If
        Next k1
        
        
        
        '更正赔2的初始值的“比较”栏目的值
        If dataSheet.Cells(j, lose2Col) <> "" And dataSheet.Cells(j, lose2Col + 1) <> "" And dataSheet.Cells(j, lose2Col + 2) <> "" Then
            dataSheet.Cells(j, lose2Col + 3) = 横向比较(dataSheet, j, lose2Col + 3, lose2Col - LOSE1col, "A", 1)
        End If
        
        '计算bf1的初始值的"比较"栏数据
        
        If dataSheet.Cells(j, BF1Col) <> "" And dataSheet.Cells(j, BF1Col + 1) <> "" And dataSheet.Cells(j, BF1Col + 2) <> "" Then
            dataSheet.Cells(j, BF1Col + 5) = 横向比较(dataSheet, j, BF1Col + 5, 2, "D", 2)
        End If
        
        '计算bf2的初始值的"比较"栏数据
        If dataSheet.Cells(j, BF2col) <> "" And dataSheet.Cells(j, BF2col + 1) <> "" And dataSheet.Cells(j, BF2col + 2) <> "" Then
            dataSheet.Cells(j, BF2col + 5) = 横向比较(dataSheet, j, BF2col + 5, 2, "D", 2)
        End If
        
        '计算bf3的初始值的"比较"栏数据
        If dataSheet.Cells(j, BF3col) <> "" And dataSheet.Cells(j, BF3col + 1) <> "" And dataSheet.Cells(j, BF3col + 2) <> "" Then
            dataSheet.Cells(j, BF3col + 5) = 横向比较(dataSheet, j, BF3col + 5, 2, "D", 2)
        End If
        
        '重新计算方差的“标识”数据
        If dataSheet.Cells(j, varCol) <> "" And dataSheet.Cells(j, varCol + 1) <> "" And dataSheet.Cells(j, varCol + 2) <> "" Then
            dataSheet.Cells(j, varCol + 3) = 标识(dataSheet.Cells(j, varCol), dataSheet.Cells(j, varCol + 1), dataSheet.Cells(j, varCol + 2)) + 固定值比较(dataSheet.Cells(j, varCol), dataSheet.Cells(j, varCol + 1), dataSheet.Cells(j, varCol + 2), 1, "D")
        End If
        If dataSheet.Cells(j + 1, varCol) <> "" And dataSheet.Cells(j + 1, varCol + 1) <> "" And dataSheet.Cells(j + 1, varCol + 2) <> "" Then
            dataSheet.Cells(j + 1, varCol + 3) = 标识(dataSheet.Cells(j + 1, varCol), dataSheet.Cells(j + 1, varCol + 1), dataSheet.Cells(j + 1, varCol + 2)) + 固定值比较(dataSheet.Cells(j + 1, varCol), dataSheet.Cells(j + 1, varCol + 1), dataSheet.Cells(j + 1, varCol + 2), 1, "D")
        End If
        If dataSheet.Cells(j + 2, varCol) <> "" And dataSheet.Cells(j + 2, varCol + 1) <> "" And dataSheet.Cells(j + 2, varCol + 2) <> "" Then
            dataSheet.Cells(j + 2, varCol + 3) = 标识(dataSheet.Cells(j + 2, varCol), dataSheet.Cells(j + 2, varCol + 1), dataSheet.Cells(j + 2, varCol + 2)) + 固定值比较(dataSheet.Cells(j + 2, varCol), dataSheet.Cells(j + 2, varCol + 1), dataSheet.Cells(j + 2, varCol + 2), 1, "D")
        End If
    Else
        '处理即时值1和即时值2中对应的“标识”列和“比较”列
        For k1 = 2 To UBound(configData)
            paraCnt = CInt(configData(k1, 9))     '基础数据个数
            dataBeginCol = CInt(dataColDict.Item(configData(k1, 2)))    '数据开始位置
            paraCompType = CStr(configData(k1, 8))            '比较数据类型
            
            '处理数据的颜色
            For j1 = 0 To paraCnt - 1
                 '颜色区分
                If dataSheet.Cells(j - 1, dataBeginCol + j1) > dataSheet.Cells(j, dataBeginCol + j1) Then
                    dataSheet.Cells(j, dataBeginCol + j1).Font.Color = downColor
                ElseIf dataSheet.Cells(j - 1, dataBeginCol + j1) < dataSheet.Cells(j, dataBeginCol + j1) Then
                    dataSheet.Cells(j, dataBeginCol + j1).Font.Color = upColor
                Else
                    dataSheet.Cells(j, dataBeginCol + j1).Font.Color = vbBlack
                End If
            Next j1
            
            If CBool(configData(k1, 7)) Then          '是否具有标识列配置项
                dataSheet.Cells(j, dataBeginCol + paraCnt) = 标识(dataSheet.Cells(j, dataBeginCol), dataSheet.Cells(j, dataBeginCol + 1), dataSheet.Cells(j, dataBeginCol + 2))
                paraCnt = paraCnt + 1
            End If
            
            '处理比较数据列
            If configData(k1, 8) = "A" Or configData(k1, 8) = "D" Then
                Call 比较(dataSheet, j, dataBeginCol + paraCnt, paraCnt, paraCompType)
            End If
        Next k1
        
         '重新计算方差的“标识”数据
        If dataSheet.Cells(j, varCol) <> "" And dataSheet.Cells(j, varCol + 1) <> "" And dataSheet.Cells(j, varCol + 2) <> "" Then
            dataSheet.Cells(j, varCol + 3) = 标识(dataSheet.Cells(j, varCol), dataSheet.Cells(j, varCol + 1), dataSheet.Cells(j, varCol + 2)) + 固定值比较(dataSheet.Cells(j, varCol), dataSheet.Cells(j, varCol + 1), dataSheet.Cells(j, varCol + 2), 1, "D")
        End If
            
    End If
Next


MsgBox ("手工数据刷新完毕！")

End Sub



Sub 显示筛选全集数据()

Dim row As Long
Dim col As Integer
Dim cc As Integer

Dim dataSheet As Worksheet

Dim i, j, k
Dim Loc As Long
Dim isOk1 As Boolean
Dim isOk2 As Boolean
Dim isOk As Boolean

Dim data()
Dim data1()

Set dataSheet = ActiveWorkbook.Sheets("综合数据")

row = dataSheet.UsedRange.Rows(dataSheet.UsedRange.Rows.Count).row
col = dataSheet.UsedRange.Columns(dataSheet.UsedRange.Columns.Count).Column



ReDim data(row - 3, 2)
Loc = 0
For i = 4 To row
    If dataSheet.Cells(i, 1).EntireRow.Hidden = False Then
        data(Loc, 0) = dataSheet.Cells(i, 9)   '编号
        data(Loc, 1) = i
        data(Loc, 2) = dataSheet.Cells(i, 6) '数据类型
        Loc = Loc + 1
    End If
Next i

ReDim data1(Loc - 1)
If Loc > 0 Then
    data1(0) = CStr(data(0, 0))
    k = 1
    For i = 1 To Loc - 1
        If data(i, 0) <> data(i - 1, 0) Then
            data1(k) = CStr(data(i, 0))
            k = k + 1
        End If
    Next i
    ReDim Preserve data1(k - 1)
End If

dataSheet.Range(Cells(3, 1), Cells(row, col)).AutoFilter
dataSheet.Range(Cells(3, 1), Cells(row, col)).AutoFilter Field:=9, Criteria1:=data1, Operator:=xlFilterValues

Set dataSheet = Nothing
ReDim data(0)
ReDim data1(0)
End Sub





Sub 记录盘口即时值(dataSheet As Worksheet, dataW, Loc, dataWcol, i, initialBgCol As Integer, bgCol As Integer, Optional isLabel As Boolean = False, Optional compareType As String = "A", Optional cnt As Integer = 3)
'------------------------------------------------------------
'  create by  ljqu   2015.3.29
'  dataSheet 待记录的工作表
'  dataW 数据数组
'  loc  工作表的位置
'  dataWCol 数据存放开始列
'  i   数据数组的指针
'  initialBgCol:数据数组中要存放的即时候数据的起始列号，     add by ljqu 2015.3.29,   增加这一列便于处理初始数据和实时数据的存放不是固定的排列顺序，原来默认即时数据跟在实始数据之后
'            变更后，将原来默认的即时值坐标由bgCol -cnt 改为initialBgCol
'
'  bgCol： 数据数组中要何存的数据的起始列号,     球探网的数据是从第10列开始的

'  isLabel：是否记录标识值，True:记录，false:不记录，对于赔率数据只有比较数据，默认为false，对于凯利数据则设为true
'  CompareType：是否记录比较值，"A":按升序比较，"D":按降序
'  cnt：要过录的数据列数，默认3列，BF1、BF2、BF3是四列
'------------------------------------------------------------
Dim upColor
Dim downColor
Dim j, j1
Dim formulaStr As String

Dim pan24 As Integer    '24小时盘口位置
Dim pan8 As Integer     '8小时盘口位置
Dim panPos As Integer    '即时值一位置

Dim n1, n2, n3
'upColor = Dict.Item("REAL_UPCOLOR")
'downColor = Dict.Item("REAL_DOWNCOLOR")


'初始24小时盘口和8小时盘口位置
pan24 = initialBgCol + cnt
pan8 = initialBgCol + 2 * cnt

upColor = vbRed
downColor = vbGreen
        
    '如果初始值为空，也同时更新初始值
    If dataSheet.Cells(Loc - 2, dataWcol) = "" Or dataSheet.Cells(Loc - 2, dataWcol + 1) = "" Or dataSheet.Cells(Loc - 2, dataWcol + 2) = "" Then
    
        '过录初始值
        For j1 = 0 To cnt - 1
            dataSheet.Cells(Loc - 2, dataWcol + j1) = dataW(i, initialBgCol + j1)
            dataSheet.Cells(Loc - 2, dataWcol + j1).Font.Color = vbBlack
        Next
        j = cnt
        If isLabel Then
            'dataSheet.Cells(loc - 2, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
            dataSheet.Cells(Loc - 2, dataWcol + j) = 标识(dataSheet.Cells(Loc - 2, dataWcol), dataSheet.Cells(Loc - 2, dataWcol + 1), dataSheet.Cells(Loc - 2, dataWcol + 2))
            j = j + 1
        End If
    
        
    End If
    
    
    
    
     '--即时值1为空，则将即时值填入即时值1和即时值2
    
    
    '过录即时值一
    If dataW(i, pan8) <> "" Then
        panPos = pan8
    ElseIf dataW(i, pan24) <> "" Then
        panPos = pan24
    Else
        panPos = -1
    End If
    If panPos > 0 Then
        For j1 = 0 To cnt - 1
            dataSheet.Cells(Loc - 1, dataWcol + j1) = dataW(i, panPos + j1)
            
             '颜色区分
            If dataSheet.Cells(Loc - 2, dataWcol + j1) > dataSheet.Cells(Loc - 1, dataWcol + j1) Then
                dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = downColor
            ElseIf dataSheet.Cells(Loc - 2, dataWcol + j1) < dataSheet.Cells(Loc - 1, dataWcol + j1) Then
                dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = upColor
            Else
                dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = vbBlack
            End If
            
        Next
        
        j = cnt
        If isLabel Then
            dataSheet.Cells(Loc - 1, dataWcol + j) = 标识(dataSheet.Cells(Loc - 1, dataWcol), dataSheet.Cells(Loc - 1, dataWcol + 1), dataSheet.Cells(Loc - 1, dataWcol + 2))
            j = j + 1
        End If
        
        '计算比较列
        If compareType = "A" Or compareType = "D" Then
            Call 比较(dataSheet, Loc - 1, dataWcol + j, j, compareType)
        End If
    End If   'end if panPos
    
    
        
    '记录即时值2
        
    For j1 = 0 To cnt - 1
        dataSheet.Cells(Loc, dataWcol + j1) = dataW(i, bgCol + j1)
        
         '颜色区分——即时值二：即时值
        If dataSheet.Cells(Loc - 1, dataWcol + j1) > dataSheet.Cells(Loc, dataWcol + j1) Then
            dataSheet.Cells(Loc, dataWcol + j1).Font.Color = downColor
        ElseIf dataSheet.Cells(Loc - 1, dataWcol + j1) < dataSheet.Cells(Loc, dataWcol + j1) Then
            dataSheet.Cells(Loc, dataWcol + j1).Font.Color = upColor
        Else
            dataSheet.Cells(Loc, dataWcol + j1).Font.Color = vbBlack
        End If
    Next


    '------即时值二
    j = cnt
    If isLabel Then
        dataSheet.Cells(Loc, dataWcol + j) = 标识(dataSheet.Cells(Loc, dataWcol), dataSheet.Cells(Loc, dataWcol + 1), dataSheet.Cells(Loc, dataWcol + 2))
        j = j + 1
    End If
    '计算比较列
    If compareType = "A" Or compareType = "D" Then
        Call 比较(dataSheet, Loc, dataWcol + j, j, compareType)
    End If

End Sub


Sub 记录联赛积分即时值(dataSheet As Worksheet, dataW, Loc, dataWcol, i, initialBgCol As Integer, midCol As Integer, bgCol As Integer, Optional isLabel As Boolean = False, Optional compareType As String = "A", Optional cnt As Integer = 3)
'------------------------------------------------------------
'  create by  ljqu   2015.3.29
'  dataSheet 待记录的工作表
'  dataW 数据数组
'  loc  工作表的位置
'  dataWCol 数据存放开始列
'  i   数据数组的指针
'  initialBgCol:数据数组中要存放的即时候数据的起始列号，     add by ljqu 2015.3.29,   增加这一列便于处理初始数据和实时数据的存放不是固定的排列顺序，原来默认即时数据跟在实始数据之后
'            变更后，将原来默认的即时值坐标由bgCol -cnt 改为initialBgCol
'
'  bgCol： 数据数组中要何存的数据的起始列号,     球探网的数据是从第10列开始的

'  isLabel：是否记录标识值，True:记录，false:不记录，对于赔率数据只有比较数据，默认为false，对于凯利数据则设为true
'  CompareType：是否记录比较值，"A":按升序比较，"D":按降序
'  cnt：要过录的数据列数，默认3列，BF1、BF2、BF3是四列
'------------------------------------------------------------
Dim upColor
Dim downColor
Dim j, j1
Dim formulaStr As String

Dim panPos As Integer    '即时值一位置

Dim n1, n2, n3

'初始24小时盘口和8小时盘口位置

upColor = vbRed
downColor = vbGreen
        

    '过录初始值
    For j1 = 0 To cnt - 1
        dataSheet.Cells(Loc - 2, dataWcol + j1) = dataW(i, initialBgCol + j1)
        'dataSheet.Cells(loc - 2, dataWcol + j1).Font.Color = vbBlack
    Next
    j = cnt
    If isLabel Then
        'dataSheet.Cells(loc - 2, dataWcol + j).FormulaR1C1 = "=标识(RC[-3],RC[-2],RC[-1])"
        dataSheet.Cells(Loc - 2, dataWcol + j) = 标识(dataSheet.Cells(Loc - 2, dataWcol), dataSheet.Cells(Loc - 2, dataWcol + 1), dataSheet.Cells(Loc - 2, dataWcol + 2))
        j = j + 1
    End If
    
    
     '--即时值1为空，则将即时值填入即时值1和即时值2
    
    
    '过录即时值一

    panPos = midCol

    For j1 = 0 To cnt - 1
        dataSheet.Cells(Loc - 1, dataWcol + j1) = dataW(i, panPos + j1)
    Next
    
    j = cnt
    If isLabel Then
        dataSheet.Cells(Loc - 1, dataWcol + j) = 标识(dataSheet.Cells(Loc - 1, dataWcol), dataSheet.Cells(Loc - 1, dataWcol + 1), dataSheet.Cells(Loc - 1, dataWcol + 2))
        j = j + 1
    End If
    
    '计算比较列
    If compareType = "A" Or compareType = "D" Then
        Call 比较(dataSheet, Loc - 1, dataWcol + j, j, compareType)
    End If

    
    
        
    '记录即时值2
        
    For j1 = 0 To cnt - 1
        dataSheet.Cells(Loc, dataWcol + j1) = dataW(i, bgCol + j1)
    Next


    '------即时值二
    j = cnt
    If isLabel Then
        dataSheet.Cells(Loc, dataWcol + j) = 标识(dataSheet.Cells(Loc, dataWcol), dataSheet.Cells(Loc, dataWcol + 1), dataSheet.Cells(Loc, dataWcol + 2))
        j = j + 1
    End If
    '计算比较列
    If compareType = "A" Or compareType = "D" Then
        Call 比较(dataSheet, Loc, dataWcol + j, j, compareType)
    End If

End Sub


Sub 计算盘形分析值(dataSheet As Worksheet, Loc)
'计算主队+客队盘形分析（绝对数据）+ 主队+客队盘形分析（相对数据)

'当客方数据没有时，原来的程序会出现类型 不匹配的错误，20150718，后改为先判断

Dim anaabsoCol
Dim anaratioCol
Dim scoremCol
Dim scoresCol
Dim anaratio1Col
Dim i
Dim sum1 As Long

Dim a1
Dim b1


anaabsoCol = dataColDict.Item("ANAABSO")
anaratioCol = dataColDict.Item("ANARATIO")
anaratio1Col = dataColDict.Item("ANARATIO_1")

scoremCol = dataColDict.Item("SCOREM")
scoresCol = dataColDict.Item("SCORES")

'先处理绝对数据
For i = 0 To 2
    
    
    If dataSheet.Cells(Loc - i, scoremCol) = "" Or dataSheet.Cells(Loc - i, scoremCol) = " " Then
        a1 = 0
    Else
        a1 = dataSheet.Cells(Loc - i, scoremCol)
    End If
    
    If dataSheet.Cells(Loc - i, scoresCol + 2) = "" Or dataSheet.Cells(Loc - i, scoresCol + 2) = " " Then
        b1 = 0
    Else
        b1 = dataSheet.Cells(Loc - i, scoresCol + 2)
    End If
    dataSheet.Cells(Loc - i, anaabsoCol) = a1 + b1
    
    
    If dataSheet.Cells(Loc - i, scoremCol + 1) = "" Or dataSheet.Cells(Loc - i, scoremCol + 1) = " " Then
        a1 = 0
    Else
        a1 = dataSheet.Cells(Loc - i, scoremCol + 1)
    End If
    
    If dataSheet.Cells(Loc - i, scoresCol + 1) = "" Or dataSheet.Cells(Loc - i, scoresCol + 1) = " " Then
        b1 = 0
    Else
        b1 = dataSheet.Cells(Loc - i, scoresCol + 1)
    End If
    
    
    dataSheet.Cells(Loc - i, anaabsoCol + 1) = a1 + b1  'dataSheet.Cells(loc - i, scoremCol + 1) + dataSheet.Cells(loc - i, scoresCol + 1)
    
    
    If dataSheet.Cells(Loc - i, scoremCol + 2) = "" Or dataSheet.Cells(Loc - i, scoremCol + 2) = " " Then
        a1 = 0
    Else
        a1 = dataSheet.Cells(Loc - i, scoremCol + 2)
    End If
    
    If dataSheet.Cells(Loc - i, scoresCol) = "" Or dataSheet.Cells(Loc - i, scoresCol) = " " Then
        b1 = 0
    Else
        b1 = dataSheet.Cells(Loc - i, scoresCol)
    End If
    
    dataSheet.Cells(Loc - i, anaabsoCol + 2) = a1 + b1 'dataSheet.Cells(loc - i, scoremCol + 2) + dataSheet.Cells(loc - i, scoresCol)
    sum1 = CLng(dataSheet.Cells(Loc - i, anaabsoCol) + dataSheet.Cells(Loc - i, anaabsoCol + 1) + dataSheet.Cells(Loc - i, anaabsoCol + 2))
    If sum1 <> 0 Then
        dataSheet.Cells(Loc - i, anaratioCol) = dataSheet.Cells(Loc - i, anaabsoCol) / sum1
        dataSheet.Cells(Loc - i, anaratioCol + 1) = dataSheet.Cells(Loc - i, anaabsoCol + 1) / sum1
        dataSheet.Cells(Loc - i, anaratioCol + 2) = dataSheet.Cells(Loc - i, anaabsoCol + 2) / sum1
    End If
    dataSheet.Cells(Loc - i, anaratioCol + 3) = dataSheet.Cells(Loc - i, anaabsoCol) * 3 + dataSheet.Cells(Loc - i, anaabsoCol + 1) * 1
    
    
Next

For i = 0 To 2
    'add by ljqu 2018.7.15
    dataSheet.Cells(Loc - i, anaratio1Col) = dataSheet.Cells(Loc - 2, anaratioCol + 3)
    dataSheet.Cells(Loc - i, anaratio1Col + 1) = dataSheet.Cells(Loc - 1, anaratioCol + 3)
    dataSheet.Cells(Loc - i, anaratio1Col + 2) = dataSheet.Cells(Loc, anaratioCol + 3)
Next

End Sub


Sub 排名分析(sheet1 As Worksheet, Loc, dataColDict As Object, teamClassDict As Object, classStrDict As Object, league)
'--------------------------------------------------------------------------------------------------------------
'参数说明：
'    sheet1:对应的综合数据sheet页
'    loc:   要修改的数据的指针,loc为即时值二对应的行号，而此处总排名在原始值上，因此需要记录在loc-2的行上
'    dataColDict：要修改的数据项对应的列字典
'    teamClassDict：联赛与对应分级类型字典，对应【01赛事】页中的“球探网名称”和 "TeamClass"列
'    classStrDict： 排名与等级对应字典，对应【TeamClass】页中的“等级分类编号”和 “等级分类串”
'    league：联赛名称
'--------------------------------------------------------------------------------------------------------------
Dim priSeq As Integer     '主队总排名
Dim secSeq As Integer     '客队总排名
Dim priBgCol As Long      '主队总排名数据所在列
Dim secBgCol As Long        '客队总排名数据所在列
Dim tmClBgCol As Long       '总排名交锋等级数据所在列
Dim teamClass As String     '联赛对应的分类类型
Dim classStr As String      '分类类型 对应的排名字串
Dim priClass As String      '主队排名对应的分类
Dim secClass As String      '客队排名对应的分类
Dim diffClass As Integer    '级差
Dim scoreClass As String    '主客队对应的分类字串

priBgCol = dataColDict.Item("SCOREM")
secBgCol = dataColDict.Item("SCORES")
tmClBgCol = dataColDict.Item("TEAMCLS")


If IsNumeric(sheet1.Cells(Loc - 2, priBgCol + 4)) Then
    priSeq = sheet1.Cells(Loc - 2, priBgCol + 4)
Else
    priSeq = 0
End If

If IsNumeric(sheet1.Cells(Loc - 2, secBgCol + 4)) Then
    secSeq = sheet1.Cells(Loc - 2, secBgCol + 4)
Else
    secSeq = 0
End If

If priSeq > 0 And secSeq > 0 Then
    teamClass = teamClassDict.Item(league)
    classStr = classStrDict.Item(teamClass)
    
    priClass = Mid(classStr, priSeq, 1)
    secClass = Mid(classStr, secSeq, 1)
    
    If Abs(priSeq - secSeq) = 1 Then     '排名连续判断
        If priSeq < secSeq Then
            scoreClass = priClass & priClass
            diffClass = 0
        Else
            scoreClass = "-" & priClass & priClass
            diffClass = 0
        End If
    Else    '不连续
        If priClass <> secClass Then  '两个分类是否相同判断
            scoreClass = priClass & secClass
            diffClass = 0 - (Asc(priClass) - Asc(secClass))
        Else   '相等
            If priSeq < secSeq Then
                scoreClass = priClass & secClass
                diffClass = 0
            Else
                scoreClass = "-" & priClass & secClass
                diffClass = 0
            End If
        
        End If   '两个分类是否相同判断
        
    End If    '排名连续判断
    
    '更新数据表中的数据
    sheet1.Cells(Loc - 2, tmClBgCol) = scoreClass
    sheet1.Cells(Loc - 2, tmClBgCol + 1) = diffClass
    
    sheet1.Cells(Loc - 1, tmClBgCol) = scoreClass
    sheet1.Cells(Loc - 1, tmClBgCol + 1) = diffClass
    
    sheet1.Cells(Loc, tmClBgCol) = scoreClass
    sheet1.Cells(Loc, tmClBgCol + 1) = diffClass
End If

End Sub




Sub 主客队排名分析(sheet1 As Worksheet, Loc, dataColDict As Object, teamClassDict As Object, classStrDict As Object, league)
'--------------------------------------------------------------------------------------------------------------
'参数说明：
'    sheet1:对应的综合数据sheet页
'    loc:   要修改的数据的指针,loc为即时值二对应的行号，而此处总排名在原始值上，因此需要记录在loc-2的行上
'    dataColDict：要修改的数据项对应的列字典
'    teamClassDict：联赛与对应分级类型字典，对应【01赛事】页中的“球探网名称”和 "TeamClass"列
'    classStrDict： 排名与等级对应字典，对应【TeamClass】页中的“等级分类编号”和 “等级分类串”
'    league：联赛名称
'--------------------------------------------------------------------------------------------------------------
Dim priSeq As Integer     '主队总排名
Dim secSeq As Integer     '客队总排名
Dim priBgCol As Long      '主队总排名数据所在列
Dim secBgCol As Long        '客队总排名数据所在列
Dim tmClBgCol As Long       '总排名交锋等级数据所在列
Dim teamClass As String     '联赛对应的分类类型
Dim classStr As String      '分类类型 对应的排名字串
Dim priClass As String      '主队排名对应的分类
Dim secClass As String      '客队排名对应的分类
Dim diffClass As Integer    '级差
Dim scoreClass As String    '主客队对应的分类字串

priBgCol = dataColDict.Item("SCOREM")
secBgCol = dataColDict.Item("SCORES")
tmClBgCol = dataColDict.Item("MSTMCLS")


If IsNumeric(sheet1.Cells(Loc - 1, priBgCol + 4)) Then
    priSeq = sheet1.Cells(Loc - 1, priBgCol + 4)
Else
    priSeq = 0
End If

If IsNumeric(sheet1.Cells(Loc - 1, secBgCol + 4)) Then
    secSeq = sheet1.Cells(Loc - 1, secBgCol + 4)
Else
    secSeq = 0
End If

If priSeq > 0 And secSeq > 0 Then
    teamClass = teamClassDict.Item(league)
    classStr = classStrDict.Item(teamClass)
    
    priClass = Mid(classStr, priSeq, 1)
    secClass = Mid(classStr, secSeq, 1)
    
    If Abs(priSeq - secSeq) = 1 Then     '排名连续判断
        If priSeq < secSeq Then
            scoreClass = priClass & priClass
            diffClass = 0
        Else
            scoreClass = "-" & priClass & priClass
            diffClass = 0
        End If
    Else    '不连续
        If priClass <> secClass Then  '两个分类是否相同判断
            scoreClass = priClass & secClass
            diffClass = 0 - (Asc(priClass) - Asc(secClass))
        Else   '相等
            If priSeq < secSeq Then
                scoreClass = priClass & secClass
                diffClass = 0
            Else
                scoreClass = "-" & priClass & secClass
                diffClass = 0
            End If
        
        End If   '两个分类是否相同判断
        
    End If    '排名连续判断
    
    '更新数据表中的数据
    sheet1.Cells(Loc - 2, tmClBgCol) = scoreClass
    sheet1.Cells(Loc - 2, tmClBgCol + 1) = diffClass
    
    sheet1.Cells(Loc - 1, tmClBgCol) = scoreClass
    sheet1.Cells(Loc - 1, tmClBgCol + 1) = diffClass
    
    sheet1.Cells(Loc, tmClBgCol) = scoreClass
    sheet1.Cells(Loc, tmClBgCol + 1) = diffClass
End If

End Sub



Sub 方差分析及记录(dataSheet As Worksheet, Loc, dataWcol, initialBgCol As Integer, bgCol As Integer, Optional isLabel As Boolean = False, Optional compareType As String = "A", Optional cnt As Integer = 3)
'------------------------------------------------------------
'  dataSheet 待记录的工作表
'  dataW 数据数组
'  loc  工作表的位置
'  dataWCol 数据存放开始列
'  i   数据数组的指针
'  initialBgCol:数据数组中要存放的即时候数据的起始列号，     add by ljqu 2015.3.29,   增加这一列便于处理初始数据和实时数据的存放不是固定的排列顺序，原来默认即时数据跟在实始数据之后
'            变更后，将原来默认的即时值坐标由bgCol -cnt 改为initialBgCol
'
'  bgCol： 数据数组中要何存的数据的起始列号,     球探网的数据是从第10列开始的

'  isLabel：是否记录标识值，True:记录，false:不记录，对于赔率数据只有比较数据，默认为false，对于凯利数据则设为true
'  CompareType：是否记录比较值，"A":按升序比较，"D":按降序
'  cnt：要过录的数据列数，默认3列，BF1、BF2、BF3是四列
'------------------------------------------------------------
Dim upColor
Dim downColor
Dim j, j1
Dim formulaStr As String

Dim n1, n2, n3
Dim dataW
Dim num1, num2

ReDim dataW(cnt)

upColor = vbRed
downColor = vbGreen
    If dataSheet.Cells(Loc, initialBgCol) <> "" And dataSheet.Cells(Loc, initialBgCol + 1) <> "" And dataSheet.Cells(Loc, initialBgCol + 2) <> "" Then
            
        For j1 = 0 To cnt - 1
                '计算值
                If IsNumeric(dataSheet.Cells(Loc - 1, initialBgCol + j1)) Then
                    num1 = dataSheet.Cells(Loc - 1, initialBgCol + j1)
                Else
                    num1 = 0
                End If
                
                If IsNumeric(dataSheet.Cells(Loc - 1, initialBgCol + j1)) Then
                    num2 = dataSheet.Cells(Loc, initialBgCol + j1)
                Else
                    num2 = 0
                End If
                
                If num1 <> 0 Or num2 <> 0 Then
                        dataW(bgCol + j1) = Abs(Round((num1 - num2) / (num1 + num2), 2))
                End If
                
            If (dataSheet.Cells(Loc, dataWcol + j1).Value <> dataW(bgCol + j1)) Or dataSheet.Cells(Loc, dataWcol + j1) = "" Then
                
                  '如若即时值有改变，先判断即时值一是否需要顺序上移
                If dataSheet.Cells(Loc - 1, dataWcol + j1) <> dataSheet.Cells(Loc - 2, dataWcol + j1) Then
                    dataSheet.Cells(Loc - 2, dataWcol + j1) = dataSheet.Cells(Loc - 1, dataWcol + j1)
                End If
                
                dataSheet.Cells(Loc - 1, dataWcol + j1) = dataSheet.Cells(Loc, dataWcol + j1)
                dataSheet.Cells(Loc, dataWcol + j1) = dataW(bgCol + j1)
            End If
            
            
             '颜色区分---即时值一：初始值
            If dataSheet.Cells(Loc - 2, dataWcol + j1) > dataSheet.Cells(Loc - 1, dataWcol + j1) Then
                dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = downColor
            ElseIf dataSheet.Cells(Loc - 2, dataWcol + j1) < dataSheet.Cells(Loc - 1, dataWcol + j1) Then
                dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = upColor
            Else
                dataSheet.Cells(Loc - 1, dataWcol + j1).Font.Color = vbBlack
            End If
            
             '颜色区分——即时值二：即时值
            If dataSheet.Cells(Loc - 1, dataWcol + j1) > dataSheet.Cells(Loc, dataWcol + j1) Then
                dataSheet.Cells(Loc, dataWcol + j1).Font.Color = downColor
            ElseIf dataSheet.Cells(Loc - 1, dataWcol + j1) < dataSheet.Cells(Loc, dataWcol + j1) Then
                dataSheet.Cells(Loc, dataWcol + j1).Font.Color = upColor
            Else
                dataSheet.Cells(Loc, dataWcol + j1).Font.Color = vbBlack
            End If
        Next
        
        '计算标识和比较列
        '------即时值一
        j = cnt
        If isLabel Then
            dataSheet.Cells(Loc - 1, dataWcol + j) = 标识(dataSheet.Cells(Loc - 1, dataWcol), dataSheet.Cells(Loc - 1, dataWcol + 1), dataSheet.Cells(Loc - 1, dataWcol + 2))
            j = j + 1
        End If
        
        '计算比较列
        If compareType = "A" Or compareType = "D" Then
            Call 比较(dataSheet, Loc - 1, dataWcol + j, j, compareType)
        End If
        

        '------即时值二
        j = cnt
        If isLabel Then
            dataSheet.Cells(Loc, dataWcol + j) = 标识(dataSheet.Cells(Loc, dataWcol), dataSheet.Cells(Loc, dataWcol + 1), dataSheet.Cells(Loc, dataWcol + 2))
            j = j + 1
        End If
        '计算比较列
        If compareType = "A" Or compareType = "D" Then
            Call 比较(dataSheet, Loc, dataWcol + j, j, compareType)
        End If
        
    End If

End Sub


Sub 计算必发指数(dataColDict, dataSheet, Loc)
''
'' 计算必发指数
'' dataColDict为列位置字典
'' dataSheet为数据页
'' Loc为数据填充的行号
''
Dim bfzsCol
Dim priceCol
Dim volCol
Dim total
Dim i
Dim sum1, sum2, sum3
Dim zs1, zs2, zs3

If Not dataColDict.exists("BFZS") Or Not dataColDict.exists("BFVOL") Or Not dataColDict.exists("BFPRICE") Then
    Exit Sub
End If
bfzsCol = dataColDict("BFZS")
volCol = dataColDict("BFVOL")
priceCol = dataColDict("BFPRICE")

For i = 0 To 2
    sum1 = dataSheet.Cells(Loc - i, priceCol) * dataSheet.Cells(Loc - i, volCol)   '胜
    sum2 = dataSheet.Cells(Loc - i, priceCol + 1) * dataSheet.Cells(Loc - i, volCol + 1) '平
    sum3 = dataSheet.Cells(Loc - i, priceCol + 2) * dataSheet.Cells(Loc - i, volCol + 2) '负
    total = sum1 + sum2 + sum3
    If total > 0 Then
        zs1 = Round(sum1 / total, 4)
        zs2 = Round(sum2 / total, 4)
        zs3 = Round(sum3 / total, 4)
        dataSheet.Cells(Loc - i, bfzsCol) = zs1
        dataSheet.Cells(Loc - i, bfzsCol + 1) = zs2
        dataSheet.Cells(Loc - i, bfzsCol + 2) = zs3
    End If
Next

End Sub
