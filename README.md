﻿# lottery
vba scripts<br>

## Change Log
### 2018-3-18 
"将模式四、模式五、模式七、模式八中的即时值一与初始值进行相减，记录数值为正的列，胜为正，则标记为""3"",平为正，则标记为""1"";负为正，则标记为""0""。同理，标记即时值二与即时值一。将模式七与模式四、模式八与模式五进行逐行比较，并进行标记。
<br>
#### 修改：<br>
    1.模式四、五、六、七、八均增加标识列schema4-schema8
    2.模式四、五后增加2列，名称为：模式X值，模式X比较
    3.模式七、八后增加3列，名称为：模式X值，模式X比较，四七和五八比较
    4.在userdef模块中增加函数： ConcateData,MethodCompare
    5.修改deal模块中的过程：模式计算"
	
	
### 2018-7-15
#### 修改：<br>
	1.将主队+客队盘形换算值(相对值）(ANARATIO)中的初始值、即时值一、即时值二中的盘形换算值单独列出，增加3列，标题为“初始、即时一、即时二”，列标识为“ANARATIO_1”，初始值、即时值一、即时值二均相同。
	2.将澳门盘口评测(PANM)中的实始值（初始盘口数据），即时值一(赛前8小时盘口数据)，即时值二(最新盘口），分列为3个栏目，栏标识别为"PANM_1","PANM_2","PANM_3"，数据项包括：主场贴水、盘口、客场贴水。
	3.将Bet365盘口评测(PANB)中的实始值（初始盘口数据），即时值一(赛前8小时盘口数据)，即时值二(最新盘口），分列为3个栏目，栏标识别为"PANB_1","PANB_2","PANB_3"，数据项包括：主场贴水、盘口、客场贴水。

### 2018-8-18
#### 修改：<br>
	1.将模式七和模式八的初始值、即时值一、即时值二单独列出，分别增加3列。
	2.将bet365数据，也如澳门盘口一样，增加针对bet365数据的模式6.
	3.威廉初始、即时1、即时2胜平负，增加九列。
	4.bf1初始值、即时1、即时2中的标识和比较分别单列，增加6列
	5.ok30初始值、即时1、即时2中的标识和比较分别单列，增加6列
	6.将模式一、二、四的初始值、即时值一、即时值二单独列出，分别增加3列。
	
### 2018-8-22
#### 修改:<br>
	优化模式计算过程，加快计算过程。
	
### 2018-10-18
####修改:<br>
	将澳客网的数据取数从【竞彩】部分改为【当场】取值修改：将取数网址由http://www.okooo.com/jingcai/shuju/，改为http://www.okooo.com/danchang/shuju/，修改网站历史数据中采集澳客网必发盈亏、胜负指数、评测数据、凯利指数四个过程。

### 2018-10-27
####修改<br>
	1. Bet365、澳门彩票、99家平均比例、赔1、赔2的初始、即时1、即时2胜平负，各增加九列。

### 2018-12-14
####修改<br>
	1. 修正澳客网必发盈亏、评测数据、凯利指数第2页及以后数据取不到的问题。
	2. 修正32位和64位office兼容的问题。

### 2019-1-3
####修改<br>
	1. 修正澳客网必发盈亏、评测数据、凯利指数中上一年12月的数据日期年份不对的问题，导致数据不更新。
	
### 2019-9-7
####修改<br>
    1. 对威廉、bet365、澳门三大类数据都增加返还率，每类增加4列，共计12列，
	2. 对立博、易胜博、赔一、赔二四大类数据都增加返还率，每类增加1列，共计4列，
	3. BF1的返还率填入【环球网必发】中的列Beffa中的返还率2
	4. 99家平均只留有原始数据，删除平铺开的初始值、即时值一、即时值二各三列数据，共9列，
	5. 竞彩比例的数据没法采集，故删除，共删除4列。
	
### 2019-12-2
####修改<br>
	1.由于澳客网的网页有所变化，默认显示的是不含有已结束赛事的数据，而在实际中需要取已结束的数据，因而先对第一页进行预取，获取相关的参数；然后再取含已结束赛事的第一页数据，并重新获取记录数和页码，再根据最新的页码进行数据获取。
	

### 2019-12-7
####修改<br>
    1. 对澳客网的球队名称，根据用户定义的对应关系表，转换成与球探网一致的球队名称，以便于数据更新。
