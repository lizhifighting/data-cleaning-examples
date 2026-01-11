 *设定工作路径
 cd "D:\B站\dataclean"
 
 *创建一个空白的数据文件来装处理后的文件
 clear
 set obs 0
 save "Data\Clndata\表5.dta", replace emptyok
 
 *正则表达式https://zhuanlan.zhihu.com/p/133457662
 *导入数据
 *先清理单个表，然后利用循环语句批量处理其他表
 *两个循环
 local wjjs "BNSF20141029-20170322 NS20141022-20140322 KCS20141022-20140322 CSX20141022-20170322   CP20141022-20170322 CN20141029-20170322"
 
 foreach wjj of local wjjs {
 	cd "D:\B站\dataclean\Data\Sdata\transportation\\`wjj'"
 	local files : dir . file "*.xls*", respectcase
	dis `"`files'"'
 	foreach file of local files {
		dis `"`file'"'
		import excel using "`file'", sheet("`r(worksheet_1)'") allstring clear 
		gen Category = "", before(A)
		replace A = strtrim(A)
		replace Category= usubstr(A,1,1)  if  regexm(A,"^[0-9][\.].")&strmatch(A,"*Week*")

		gen Measure = A if regexm(A,"^[0-9][\.].")&strmatch(A,"*Week*"),after(Category)
		replace Measure = subinstr(Measure,"5.","",.)
		
		gen date = ustrregexs(3) if ustrregexm("`file'","(.*)(data_)(.*)(.xls.*)")
		gen Railroad = substr("`wjj'",1,strpos("`wjj'","2014")-1),before(Category)
		drop in 1/4
		dropmiss, obs force  //删除整行缺失的行
		carryforward Category Measure, replace 
		keep if Category == "5"

		*删去多余的行
		drop if strmatch(A,"*5*")&B==""
		foreach x of varlist * {
			replace `x'="Total" if `x'[_n]==""&`x'[_n-1]=="Total"
		}
		drop if (A=="Train Type"|A==" Train Type")&(B==""|B=="Cause")
		rename A Variable

		replace G=G[2] in 1 if F[1]=="Other"&G[1]==""
		foreach x of varlist * {
			replace `x' = subinstr(`x'," ","_",.)  in 1
		}
		replace Variable = "" in 1
		replace Category = "" in 1
		replace date = "" in 1
		replace Railroad = "" in 1
		replace Measure = "" in 1
		nrow 1

		cap findname, all(missing(@))
		cap drop `r(varlist)'

		drop if Variable==""
		cap gen Other=""
		cap ren Locomotive_Power Locomotive_power
		append using "D:\B站\dataclean\Data\Clndata\表5.dta"
		save "D:\B站\dataclean\Data\Clndata\表5.dta", replace		
	}
 }


use "D:\B站\dataclean\Data\Clndata\表5.dta", clear
replace Briefly_Explain_Cause=R if R!=""
replace Briefly_Explain_Cause=V if V!=""
replace Briefly_Explain_Cause=P if P!=""
replace Briefly_Explain_Cause=Briefly_Explain_Cause_for_Other if Briefly_Explain_Cause_for_Other!="" 
drop R V  P Briefly_Explain_Cause_for_Other
replace Track_Maintenance=Track__Maintenance if Track__Maintenance!=""
replace Track_Maintenance=Track_maintenance  if Track_maintenance!=""
drop Track__Maintenance Track_maintenance
order Mechanical_Issue Track_Maintenance  Act_of_God Congestion Connecting_Carriers Other Briefly_Explain_Cause,before(Total)
format * %15s
drop if Total==""

dropmiss Crew Locomotive_power Mechanical_Issue Track_Maintenance Act_of_God Congestion Connecting_Carriers Other Briefly_Explain_Cause,obs force
**处理日期
include C:\Users\li_zhi\Desktop\From\Shan\Code\date.do   //合并同一周内的日期
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
save "D:\B站\dataclean\Data\Clndata\f5.dta",replace
