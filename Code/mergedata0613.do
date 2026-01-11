********************************************************************************
*整理思路：分-总
*1.按总表格式清理有2014-2017年现成长数据的表1，2，3，4，7，10
*2.合并各周的表5，6，8，9，11得到长数据
*3.按总表格式清理没有2014-2017年现成长数据的表5，6，8，9，11
*4.利用merge把处理好的表1-11并入总的数据

*处理小结
*1.辅助变量/行，例如表示日期的行，由于stata不支持纯日期的变量名，因此利用s和_对日期进行处理
*2.新学的命令：insobs、nrow、dropmiss
*3.文本处理，substr，subinstr，split，ustrregexm，ustrregexr，ustrregexs
********************************************************************************
global path "C:\Users\Lizhi\Desktop\From\Shan"

cd "$path\Data"

import excel using "Sdata\transportation\EP724 Consolidated Data through 2022-08-10.xlsx",clear allstring
foreach x of varlist G-KA {              
	replace `x'=subinstr(`x',"/","_",.)   in 1
	replace `x'=subinstr(`x'," ","",.)    in 1
	replace `x'="s"+`x'                   in 1
}
replace A="Railroad" in 1
replace B="Category" in 1
replace C="Sub_Category" in 1
replace F="SubVariable" in 1
nrow 1
drop in 3317
drop if Railroad=="Chicago"
replace SubVariable=subinstr(SubVariable," ","_",.)
save Clndata\all0810.dta,replace

*调整各个表以适应总体的表
**#1-Train_Speeds
insheet using "Sdata\New folder\1-Train_Speeds.csv",clear double
keep in 1/6985                        //只保留2017年以前的数据
drop year month week
tostring mph,replace force
gen sdate=date
replace sdate=subinstr(sdate,"/","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
reshape wide  mph , i(railroad commodity) string  j(sdate) 
insobs 1,before(1)                    //插入表示日期的行
foreach x of varlist *{               //把变量标签作为第一行的值
	local a: var label `x'
	replace `x'="`a'" in 1
	replace `x'=subinstr(`x',"mph","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
}
gen Category="",after(railroad)
gen SubVariable="",after(commodity)
nrow 1
rename Commodity Variable
replace Category="1"
replace Variable="Grain unit" if Variable=="Grain"&Railroad!="CN"
replace Variable="Coal unit" if Variable=="Coal"&Railroad!="CN"
replace Variable="Automotive unit" if Variable=="Automotive"&Railroad!="CN"
replace Variable="Crude oil unit" if Variable=="Crude Oil"&Railroad!="CN"
replace Variable="Ethanol unit" if Variable=="Ethanol"&Railroad!="CN"
replace Variable="Other Unit" if Variable=="Other"&Railroad=="CN"
replace Variable="Crude" if Variable=="Crude Oil"&Railroad=="CN"
save Clndata\f1.dta,replace

**#2-Rail_Terminal_Dwell_Times
insheet using "Sdata\New folder\2-Rail_Terminal_Dwell_Times.csv",clear double
gsort year month week railroad yard
keep in 1/8128
keep date railroad yard value
tostring value,replace force          //先变为字符串型变量处理
gen sdate=date
replace sdate=subinstr(sdate,"/","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
reshape wide  value , i(railroad yard) string  j(sdate) 
insobs 1,before(1)                    //插入表示日期的行
foreach x of varlist *{               //把变量标签作为第一行的值
	local a: var label `x'
	replace `x'="`a'" in 1
	replace `x'=subinstr(`x',"value","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
}
gen Category="",after(railroad)
gen SubVariable="",after(yard)
nrow 1
rename Yard Variable
replace Category="2"
replace Variable=subinstr(Variable,"，",",",.)
replace Variable=subinstr(Variable,"ALBANY","Albany",.)
replace Variable=subinstr(Variable,"BENSENVILLE","Bensenville",.)
replace Variable=subinstr(Variable,"GLENWOOD","Glenwood",.)
replace Variable=subinstr(Variable,"HARVEY","Harvey",.)
replace Variable=subinstr(Variable,"LA CROSSE","La Crosse",.)
replace Variable=subinstr(Variable,"MASON CITY","Mason City",.)
replace Variable=subinstr(Variable,"MILWAUKEE","Milwaukee",.)
replace Variable=subinstr(Variable,"NAHANT","Nahant",.)
replace Variable=subinstr(Variable,"ST PAUL","St Paul",.)
save Clndata\f2.dta,replace

**#3-Cars_On_Line
insheet using "Sdata\New folder\3-Cars_On_Line.csv",clear double
keep in 1/8001                       //只保留2017年以前的数据
keep date railroad type cars
tostring cars,replace force          //先变为字符串型变量处理
gen sdate=date
replace sdate=subinstr(sdate,"/","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
reshape wide  cars , i(railroad type) string  j(sdate) 
insobs 1,before(1)                    //插入表示日期的行
foreach x of varlist *{               //把变量标签作为第一行的值
	local a: var label `x'
	replace `x'="`a'" in 1
	replace `x'=subinstr(`x',"cars","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
}
gen Category="",after(railroad)
gen SubVariable="",after(type)
nrow 1
rename Type Variable
replace Category="3"
save Clndata\f3.dta,replace

**#4-Rail_Origin_Dwell_Times
insheet using "Sdata\New folder\4-Rail_Origin_Dwell_Times.csv",clear double
keep in 1/5207
keep date railroad commodity hours
tostring hours,replace force 
gen sdate=date
replace sdate=subinstr(sdate,"/","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
reshape wide  hours , i(railroad commodity) string  j(sdate) 
insobs 1,before(1)                    //插入表示日期的行
foreach x of varlist *{               //把变量标签作为第一行的值
	local a: var label `x'
	replace `x'="`a'" in 1
	replace `x'=subinstr(`x',"hours","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
	replace `x'=subinstr(`x',"cars","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
}
gen Category="",after(railroad)
gen SubVariable="",after(commodity)
nrow 1
rename Commodity Variable
replace Category="4"
replace Variable="Grain unit" if Variable=="Grain"
replace Variable="Coal unit" if Variable=="Coal"
replace Variable="Automotive unit" if Variable=="Automotive"
replace Variable="Crude Oil unit" if Variable=="Crude Oil"
replace Variable="Ethanol unit" if Variable=="Ethanol"
replace Variable="All Other Unit Trains" if Variable=="Other"
save Clndata\f4.dta,replace

**#5-Trains_Held_Short
use Clndata\表5,clear
dropmiss Crew Locomotive_power Mechanical_Issue Track_Maintenance Act_of_God Congestion Connecting_Carriers Other Briefly_Explain_Cause,obs force
**处理日期
include C:\Users\Lizhi\Desktop\From\Shan\Code\date.do   //合并同一周内的日期
gen sdate=date
replace sdate=subinstr(sdate,"-","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)

egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
gather Crew Locomotive_power Mechanical_Issue Track_Maintenance Act_of_God Congestion Connecting_Carriers Other  Total
replace variable="total"  if variable=="Total"
replace variable="other"  if variable=="Other"
spread sdate value
replace Category="5"
format * %15s
replace Briefly_Explain_Cause="" if variable!="other"
rename variable SubVariable
replace Variable="Coal unit" if Variable=="Coal"&Railroad=="CN"
replace Variable="Crude oil unit" if Variable=="Crude"&Railroad=="CN"
replace Variable="Ethanol unit" if Variable=="Ethanol"&Railroad=="CN"
replace Variable="Grain unit" if Variable=="Grain"&Railroad=="CN"
replace Variable="Other unit" if Variable=="Other"&Railroad=="CN"
replace Variable="Coal unit" if Variable=="Coal Unit"&Railroad=="CSX"
replace Variable="Crude oil unit" if Variable=="Crude Oil Unit"&Railroad=="CSX"
replace Variable="Crude oil unit" if Variable=="Crude Oil unit"&Railroad=="CSX"
replace Variable="Ethanol unit" if Variable=="Ethanol Unit"&Railroad=="CSX"
replace Variable="Grain unit" if Variable=="Grain Unit"&Railroad=="CSX"
replace Variable="Other unit" if Variable=="Other Unit"&Railroad=="CSX"
foreach x of varlist s2014_10_22-s2017_3_8  {
	bys Railroad Category Variable SubVariable:replace `x'=`x'[_n-1] if `x'==""
	bys Railroad Category Variable SubVariable:replace `x'=`x'[_n+1] if `x'==""
}
duplicates drop Railroad Category Variable SubVariable,force
save Clndata\f5.dta,replace

**#6-Rail_C... Greater
use Clndata\表6,clear
keep date Category Railroad Variable Measure Loaded Empty
include C:\Users\Lizhi\Desktop\From\Shan\Code\date.do   //合并同一周内的日期
gen sdate=date
replace sdate=subinstr(sdate,"-","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
gather Loaded Empty                    //先把SubVariable宽转长
rename variable SubVariable
spread sdate value                     //再把sdate长转宽
replace Category="6"
drop if strmatch(Variable,"*Railroad*")|strmatch(Variable,"*:*")|strmatch(Variable,"*EP 724*")
replace Variable="Automotive" if Variable=="AUTOMOTIVE"&Railroad=="CN"
replace Variable="Coal" if Variable=="COAL"&Railroad=="CN"
replace Variable="Crude Oil" if Variable=="CRUDE OIL"&Railroad=="CN"
replace Variable="Crude Oil" if Variable=="Crude"&Railroad=="CN"
replace Variable="Ethanol" if Variable=="ETHANOL"&Railroad=="CN"
replace Variable="Grain" if Variable=="GRAIN"&Railroad=="CN"
replace Variable="Intermodal" if Variable=="INTERMODAL"&Railroad=="CN"
replace Variable="All Other" if Variable=="Other"&Railroad=="CN"
replace Variable="Intermodal" if Variable=="Intermodal (flat cars)"&Railroad=="CSX"
replace Variable="Automotive" if Variable=="Multilevel (automotive)"&Railroad=="KCS"
foreach x of varlist s2014_10_22-s2017_3_8  {
	bys Railroad Category Variable SubVariable:replace `x'=`x'[_n-1] if `x'==""
	bys Railroad Category Variable SubVariable:replace `x'=`x'[_n+1] if `x'==""
}
duplicates drop Railroad Category Variable SubVariable,force
save Clndata\f6.dta,replace

**#7-Grain_Rail_Cars_Loaded_and_Billed
insheet using "Sdata\New folder\7-Grain_Rail_Cars_Loaded_and_Billed.csv",clear
keep in 1/9419
keep date railroad state all
tostring all,force replace
gen sdate=date
replace sdate=subinstr(sdate,"/","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
reshape wide  all , i(railroad state) string  j(sdate) 
insobs 1,before(1)                    //插入表示日期的行
foreach x of varlist *{               //把变量标签作为第一行的值
	local a: var label `x'
	replace `x'="`a'" in 1
	replace `x'=subinstr(`x',"all","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
	replace `x'=subinstr(`x',"cars","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
}
gen Category="",after(railroad)
gen Variable="All Ordering Systems",before(state)
nrow 1
rename State SubVariable
replace Category="7"
save Clndata\f7_1.dta,replace
insheet using "Sdata\New folder\7-Grain_Rail_Cars_Loaded_and_Billed.csv",clear
keep in 1/9419
keep date railroad state dedicated_or_shuttle
tostring dedicated_or_shuttle,force replace
gen sdate=date
replace sdate=subinstr(sdate,"/","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
reshape wide  dedicated_or_shuttle, i(railroad state) string  j(sdate) 
insobs 1,before(1)                    //插入表示日期的行
foreach x of varlist *{               //把变量标签作为第一行的值
	local a: var label `x'
	replace `x'="`a'" in 1
	replace `x'=subinstr(`x',"dedicated_or_shuttle","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
	replace `x'=subinstr(`x',"cars","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
}
gen Category="",after(railroad)
gen Variable="Shuttle / Dedicated Train Service",before(state)
nrow 1
rename State SubVariable
replace Category="7"
save Clndata\f7_2.dta,replace
insheet using "Sdata\New folder\7-Grain_Rail_Cars_Loaded_and_Billed.csv",clear
keep in 1/9419
keep date railroad state other
tostring other,force replace
gen sdate=date
replace sdate=subinstr(sdate,"/","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop date Year Month Day *time*
reshape wide  other, i(railroad state) string  j(sdate) 
insobs 1,before(1)                    //插入表示日期的行
foreach x of varlist *{               //把变量标签作为第一行的值
	local a: var label `x'
	replace `x'="`a'" in 1
	replace `x'=subinstr(`x',"other","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
	replace `x'=subinstr(`x',"cars","",.) in 1
	replace `x'=subinstr(`x'," ","",.)   in 1
}
gen Category="",after(railroad)
gen Variable="Other Than Shuttle / Dedicated Train Service",before(state)
nrow 1
rename State SubVariable
replace Category="7"
save Clndata\f7_3.dta,replace

**#8-Grain_Car_Order_Fulfillment_for_Manifest_Service
use Clndata\表8,clear
include C:\Users\Lizhi\Desktop\From\Shan\Code\date.do   //合并同一周内的日期
gen sdate=date
replace sdate=subinstr(sdate,"-","_",.)
split sdate,gen("time") parse("_")
destring time*,replace
gen Year=.
gen Month=.
gen Day=.
egen mintime1=min(time1)
egen maxtime1=max(time1)
egen mintime2=min(time2)
egen maxtime2=max(time2)
egen mintime3=min(time3)
egen maxtime3=max(time3)
foreach x of varlist time*{
	replace Year=`x'  if max`x'==2017
	replace Month=`x' if max`x'==12
	replace Day=`x'   if max`x'==31
}
replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
drop if SubVariable==""|SubVariable=="Sdate"|strmatch(SubVariable,"*TOTAL*")|strmatch(SubVariable,"*Total*")|strmatch(SubVariable,"*:*")
drop if SubVariable=="** SD new orders and orders filled are all RCP&E"
drop if SubVariable=="** Note 122 of the 227 orders filled in SD were for RCP&E - the balance were CP on line customers in SD."
replace a___Running_Total_Number_of_Outs=a__Running_Total_Number_of_Outst if a__Running_Total_Number_of_Outst!=""
drop a__Running_Total_Number_of_Outst  _* date Year Month Day *time* 
duplicates drop Railroad SubVariable sdate,force
gather a___Running_Total_Number_of_Outs b__Average_Number_of_Days_Late_F c__Number_of_New_Car_Orders d__Number_of_Car_Orders_Filled e_1__Number_of_Orders_Canceled_B e_2__Number_of_Orders_Canceled_B
spread sdate value
rename variable  Variable
replace Variable="Running Total Number of Outstanding Car Orders" if Variable=="a___Running_Total_Number_of_Outs"
replace Variable="Average Number of Days Late For All Outstanding Grain Car Orders" if Variable=="b__Average_Number_of_Days_Late_F"
replace Variable="Number of New Car Orders" if Variable=="c__Number_of_New_Car_Orders"
replace Variable="Number of Car Orders Filled" if Variable=="d__Number_of_Car_Orders_Filled"
replace Variable="Number of Orders Canceled By Shipper" if Variable=="e_1__Number_of_Orders_Canceled_B"
replace Variable="Number of Orders Canceled By Railroad" if Variable=="e_2__Number_of_Orders_Canceled_B"
save Clndata\f8.dta,replace

// **#9-Coal_Unit_Train_Loadings
// use Clndata\表9,clear
// cap keep date Railroad Variable  Loadings_Plan
// include C:\Users\Lizhi\Desktop\From\Shan\Code\date.do   //合并同一周内的日期
// gen sdate=date
// replace sdate=subinstr(sdate,"-","_",.)
// split sdate,gen("time") parse("_")
// destring time*,replace
// gen Year=.
// gen Month=.
// gen Day=.
// egen mintime1=min(time1)
// egen maxtime1=max(time1)
// egen mintime2=min(time2)
// egen maxtime2=max(time2)
// egen mintime3=min(time3)
// egen maxtime3=max(time3)
// foreach x of varlist time*{
// 	replace Year=`x'  if max`x'==2017
// 	replace Month=`x' if max`x'==12
// 	replace Day=`x'   if max`x'==31
// }
// replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
// drop date Year Month Day *time*
// replace Loadings_Plan="" if Loadings_Plan=="No Coal Loadings on KCSR Lines"
// reshape wide  Loadings_Plan, i(Railroad Variable) string  j(sdate) 
// insobs 1,before(1)                    //插入表示日期的行
// foreach x of varlist *{               //把变量标签作为第一行的值
// 	local a: var label `x'
// 	replace `x'="`a'" in 1
// 	replace `x'=subinstr(`x',"Loadings_Plan","",.) in 1
// 	replace `x'=subinstr(`x'," ","",.)   in 1
// 	replace `x'=subinstr(`x',"cars","",.) in 1
// 	replace `x'=subinstr(`x'," ","",.)   in 1
// }
// gen Category="",after(Railroad)
// nrow 1
// replace Category="9"
// gen SubVariable="Loadings Plan",after(Variable)
// save Clndata\f9_1.dta,replace
//
// use Clndata\表9,clear
// cap keep date Railroad Variable  Loadings_Average
// include C:\Users\Lizhi\Desktop\From\Shan\Code\date.do   //合并同一周内的日期
// gen sdate=date
// replace sdate=subinstr(sdate,"-","_",.)
// split sdate,gen("time") parse("_")
// destring time*,replace
// gen Year=.
// gen Month=.
// gen Day=.
// egen mintime1=min(time1)
// egen maxtime1=max(time1)
// egen mintime2=min(time2)
// egen maxtime2=max(time2)
// egen mintime3=min(time3)
// egen maxtime3=max(time3)
// foreach x of varlist time*{
// 	replace Year=`x'  if max`x'==2017
// 	replace Month=`x' if max`x'==12
// 	replace Day=`x'   if max`x'==31
// }
// replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
// drop date Year Month Day *time*
// replace Loadings_Average="" if Loadings_Average=="No Coal Loadings on KCSR Lines"
// reshape wide  Loadings_Average, i(Railroad Variable) string  j(sdate) 
// insobs 1,before(1)                    //插入表示日期的行
// foreach x of varlist *{               //把变量标签作为第一行的值
// 	local a: var label `x'
// 	replace `x'="`a'" in 1
// 	replace `x'=subinstr(`x',"Loadings_Average","",.) in 1
// 	replace `x'=subinstr(`x'," ","",.)   in 1
// 	replace `x'=subinstr(`x',"cars","",.) in 1
// 	replace `x'=subinstr(`x'," ","",.)   in 1
// }
// gen Category="",after(Railroad)
// nrow 1
// replace Category="9"
// gen SubVariable="Loadings Average",after(Variable)
// save Clndata\f9_2.dta,replace
//
// **#10-Grain_Shuttle_Train_Turns
// insheet using "Sdata\New folder\10-Grain_Shuttle_Train_Turns.csv",clear double
// gsort year month week railroad region
// keep in 1/1144
// keep date railroad region averageturns
// tostring averageturns,replace
// gen sdate=date
// replace sdate=subinstr(sdate,"/","_",.)
// split sdate,gen("time") parse("_")
// destring time*,replace
// gen Year=.
// gen Month=.
// gen Day=.
// egen mintime1=min(time1)
// egen maxtime1=max(time1)
// egen mintime2=min(time2)
// egen maxtime2=max(time2)
// egen mintime3=min(time3)
// egen maxtime3=max(time3)
// foreach x of varlist time*{
// 	replace Year=`x'  if max`x'==2017
// 	replace Month=`x' if max`x'==12
// 	replace Day=`x'   if max`x'==31
// }
// replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
// drop date Year Month Day *time*
// reshape wide  averageturns , i(railroad region) string  j(sdate) 
// insobs 1,before(1)                    //插入表示日期的行
// foreach x of varlist *{               //把变量标签作为第一行的值
// 	local a: var label `x'
// 	replace `x'="`a'" in 1
// 	replace `x'=subinstr(`x',"averageturns","",.) in 1
// 	replace `x'=subinstr(`x'," ","",.)   in 1
// 	replace `x'=subinstr(`x',"cars","",.) in 1
// 	replace `x'=subinstr(`x'," ","",.)   in 1
// }
// gen Category="",after(railroad)
// nrow 1
// gen Variable="Average Turns",after(Category)
// rename Region SubVariable
// replace Category="10"
// save Clndata\f10_1.dta,replace
// insheet using "Sdata\New folder\10-Grain_Shuttle_Train_Turns.csv",clear double
// gsort year month week railroad region
// keep in 1/1144
// keep date railroad region planned
// tostring planned,replace
// gen sdate=date
// replace sdate=subinstr(sdate,"/","_",.)
// split sdate,gen("time") parse("_")
// destring time*,replace
// gen Year=.
// gen Month=.
// gen Day=.
// egen mintime1=min(time1)
// egen maxtime1=max(time1)
// egen mintime2=min(time2)
// egen maxtime2=max(time2)
// egen mintime3=min(time3)
// egen maxtime3=max(time3)
// foreach x of varlist time*{
// 	replace Year=`x'  if max`x'==2017
// 	replace Month=`x' if max`x'==12
// 	replace Day=`x'   if max`x'==31
// }
// replace sdate="s"+string(Year)+"_"+string(Month)+"_"+string(Day)
// drop date Year Month Day *time*
// reshape wide  planned , i(railroad region) string  j(sdate) 
// insobs 1,before(1)                    //插入表示日期的行
// foreach x of varlist *{               //把变量标签作为第一行的值
// 	local a: var label `x'
// 	replace `x'="`a'" in 1
// 	replace `x'=subinstr(`x',"planned","",.) in 1
// 	replace `x'=subinstr(`x'," ","",.)   in 1
// 	replace `x'=subinstr(`x',"cars","",.) in 1
// 	replace `x'=subinstr(`x'," ","",.)   in 1
// }
// gen Category="",after(railroad)
// nrow 1
// gen Variable="Average Turns",after(Category)
// rename Region SubVariable
// replace Category="10"
// save Clndata\f10_2.dta,replace

// *11-Rail_Carloadings
// insheet using "Sdata\New folder\11-Rail_Carloadings.csv",clear
// drop year month week
// gen sdate=date
// replace sdate=subinstr(sdate,"/","_",.)
// drop date
// reshape wide  carloads , i(railroad commodity type) string  j(sdate) 


use Clndata\all0810.dta,clear
merge m:1 Railroad Category Variable using Clndata\f1.dta,update replace
drop _merge
merge m:1 Railroad Category Variable using Clndata\f2.dta,update replace
drop _merge
merge m:1 Railroad Category Variable using Clndata\f3.dta,update replace
drop _merge
merge m:1 Railroad Category Variable using Clndata\f4.dta,update replace
drop _merge
merge m:1 Railroad Category Variable SubVariable using Clndata\f5.dta,update replace
drop _merge
merge m:1 Railroad Category Variable SubVariable using Clndata\f6.dta,update replace
drop _merge
merge m:1 Railroad Category Variable SubVariable using Clndata\f7_1.dta,update replace
drop _merge
merge m:1 Railroad Category Variable SubVariable using Clndata\f7_2.dta,update replace
drop _merge
merge m:1 Railroad Category Variable SubVariable using Clndata\f7_3.dta,update replace
drop _merge
merge m:1 Railroad Category Variable SubVariable Measure using Clndata\f8.dta,update replace
drop _merge
// merge m:1 Railroad Category Variable SubVariable using Clndata\f9_1.dta,update replace
// drop _merge
// merge m:1 Railroad Category Variable SubVariable using Clndata\f9_2.dta,update replace
// drop _merge
// merge m:1 Railroad Category Variable SubVariable using Clndata\f10_1.dta,update replace
// drop _merge
// merge m:1 Railroad Category Variable SubVariable using Clndata\f10_2.dta,update replace
// drop _merge
export excel using "Clndata\clndata0906.xlsx", firstrow(variables) replace