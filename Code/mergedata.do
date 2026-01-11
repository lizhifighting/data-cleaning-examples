global path "C:\Users\Lizhi\Desktop\From\Shan\Data"
set trace off
*创建一个空白数据文件来装处理后的文件
clear
foreach i of numlist 1/10 {
	set obs 0
	save $path\Clndata\表`i',replace emptyok
}
*构建循环把每个文件夹内每个文件中的每个表合并成一张大表
local wjs "BNSF20141029-20170322 NS20141022-20140322 KCS20141022-20140322 CSX20141022-20170322 CP20141022-20170322 CN20141029-20170322"
foreach wj of local wjs {
	cd "$path\Sdata\transportation\\`wj'"
	local files : dir . file "*.xls*", respectcase    //获取文件夹中所有excel表的名称
	dis `"`files'"'                                   //将获取的表名存入局部宏
	foreach file in `files' {                         //在同一个文件夹内循环
		*表5
		import excel using "`file'", sheet("`r(worksheet_`i')'") allstring clear
		dis `"`file'"'
		gen Category="",before(A)
		replace Category=usubstr(A,1,1)  if regexm(A,"^[0-9].")&strmatch(A,"*Week*")
		replace Category="4"             if strmatch(A,"*4.*")&strmatch(A,"*Week*")
		replace Category="5"             if strmatch(A,"*5.*")&strmatch(A,"*Week*")
		replace Category="1"             if A=="1. Average Train Speed by Train Type and Overall System Average Train Speed reported in Miles per Hour (MPH)"
		gen Measure=A  if regexm(A,"(.*)(^[0-9].)")&strmatch(A,"*Week*")
		replace Measure=A  			  if strmatch(A,"*4.*")&strmatch(A,"*Week*")
		replace Measure=A             if strmatch(A,"*5.*")&strmatch(A,"*Week*")
		replace Measure=A  			  if A=="1. Average Train Speed by Train Type and Overall System Average Train Speed reported in Miles per Hour (MPH)"
		gen date="`file'"
		replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
		gen Railroad=substr("`wj'",1,strpos("`wj'","2014")-1) ,before(Category)
		drop in 1/4
		dropmiss, obs force  //删除整行缺失的行
		carryforward Category Measure,replace
		foreach i of numlist 5 {    
			keep if Category=="`i'"
			drop if strmatch(A,"*`i'*")&B==""
			foreach x of varlist * {
				replace `x'="Total" if `x'[_n]==""&`x'[_n-1]=="Total"
			}
			drop if (A=="Train Type"|A==" Train Type")&(B==""|B=="Cause")
			rename A Variable
			replace G=G[2]  in 1  if F[1]=="Other"&G[1]==""
			foreach x of varlist * {
				replace `x'=subinstr(`x'," ","_",.)    in 1
			}
			replace Variable=""  in 1
			replace Category=""  in 1
			replace date=""      in 1
			replace Railroad=""  in 1
			replace Measure=""   in 1   //保持这几个变量的名称不变
			nrow 1
			cap findname, all(missing(@)) 
			cap drop `r(varlist)'
			drop if Variable==""
			cap gen Other=""
			cap ren Locomotive_Power Locomotive_power
			append using $path\Clndata\表`i'
			save $path\Clndata\表`i',replace	   //保留表`i'的全部cause
		}
		*表6
		import excel using "`file'", sheet("`r(worksheet_`i')'") allstring clear
		dis `"`file'"'
		gen Category="",before(A)
		replace Category=usubstr(A,1,1)  if regexm(A,"^[0-9].")&strmatch(A,"*Week*")
		replace Category="4"             if strmatch(A,"*4.*")&strmatch(A,"*Week*")
		replace Category="5"             if strmatch(A,"*5.*")&strmatch(A,"*Week*")
		replace Category="1"             if A=="1. Average Train Speed by Train Type and Overall System Average Train Speed reported in Miles per Hour (MPH)"
		gen Measure=A  if regexm(A,"(.*)(^[0-9].)")&strmatch(A,"*Week*")
		replace Measure=A  			  if strmatch(A,"*4.*")&strmatch(A,"*Week*")
		replace Measure=A             if strmatch(A,"*5.*")&strmatch(A,"*Week*")
		replace Measure=A  			  if A=="1. Average Train Speed by Train Type and Overall System Average Train Speed reported in Miles per Hour (MPH)"
		gen date="`file'"
		replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
		gen Railroad=substr("`wj'",1,strpos("`wj'","2014")-1) ,before(Category)
		drop in 1/4
		dropmiss, obs force  //删除整行缺失的行
		carryforward Category Measure,replace
		foreach i of numlist 6 {
			keep if Category=="`i'"
			*删除无关的行和列
			drop in 1/2
			replace B=B+"1"  in 1
			replace C=C+"1"  in 1  if C[1]!=""
			replace F=F+"1"  in 1  if C[1]==""
			replace D=D+"2"  in 1  if D[1]!=""
			replace G=G+"2"  in 1  if D[1]==""
			replace E=E+"2"  in 1  if E[1]!=""
			cap replace K=K+"2"  in 1  if E[1]==""
			rename A Variable
			foreach x of varlist * {
				replace `x'=subinstr(`x'," ","_",.)    in 1
			}
			replace Variable=""  in 1
			replace Category=""  in 1
			replace date=""      in 1
			replace Railroad=""  in 1
			replace Measure=""   in 1   //保持这几个变量的名称不变
			nrow 1
			destring Loaded1 Empty1 Loaded2 Empty2,replace force
			gen Loaded=round(Loaded1+Loaded2),before(Loaded1)
			gen Empty=round(Empty1+Empty2),before(Empty1)
			tostring Loaded* Empty*,replace force
			drop if Variable==""
			append using $path\Clndata\表`i'
			save $path\Clndata\表`i',replace
		}

		*表8
		import excel using "`file'", sheet("Grain Metrics 2 (item 8)") allstring clear
		dis `"`file'"'
		gen Category=""
		replace Category=usubstr(A,1,1) if regexm(A,"^[0-9].")
		gen Measure=""
		replace Measure=A               if regexm(A,"^[0-9].")
        gen date="`file'"
		replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
		gen Railroad=substr("`wj'",1,strpos("`wj'","2014")-1) ,before(Category)
		order Category Measure date,after(Railroad)
		drop in 1/4
		dropmiss ,obs force  //删除整行缺失的行
		carryforward Category Measure,replace
		foreach i of numlist 8 {
			keep if Category=="`i'"
			drop if A==""|(strmatch(A,"*`i'*")&B=="")
			rename A SubVariable
			replace SubVariable=""  in 1
			replace Category=""     in 1
			replace date=""         in 1
			replace Railroad=""     in 1
			replace Measure=""      in 1   //保持这几个变量的名称不变
			nrow 1
			cap findname, all(missing(@)) 
			cap drop `r(varlist)'
			append using $path\Clndata\表`i'
			save $path\Clndata\表`i',replace			
		}

// 		*表9
// 		import excel using "`file'", sheet("Grain & Coal Plans (items 9-10)") allstring clear
// 		dis `"`file'"'
// 		gen Category=""
// 		order Category
// 		replace Category=usubstr(A,1,1) if strmatch(A,"*9.*")
// 		replace Category="10"           if strmatch(A,"*10.*")
// 		gen date="`file'"
// 		replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
// 		gen Railroad=substr("`wj'",1,strpos("`wj'","2014")-1) ,before(Category)
// 		dropmiss,obs force  //删除整行缺失的行
// 		carryforward Category,replace
// 		foreach i of numlist 9 {
// 			local j=`i'+1
// 			keep if Category=="`j'"
// 			*删除无关的行和列
// 			drop in 1
// 			rename A Variable
// 			drop if Variable=="10. Average Daily Coal Unit Train Loadings vs.Plan for the Reporting Week By Coal"
// 			foreach x of varlist * {
// 				replace `x'=subinstr(`x'," ","_",.)    in 1
// 			}
// 			replace Variable=""  in 1
// 			replace Category=""  in 1
// 			replace date=""      in 1
// 			replace Railroad=""  in 1
// 			nrow 1
// 			drop if Variable==""
// 			cap ren Plan Loadings_Plan
// 			cap ren Actual Loadings_Average	
// 			keep date Railroad Category Variable Loadings_Plan Loadings_Average			
// 			append using $path\Clndata\表`j'
// 			save $path\Clndata\表`j',replace
// 		}
	}
}
**UP里的文件与其他文件夹里的不一样，单独处理
cd "$path\Sdata\transportation\UP20141022-20140322"
local files : dir . file "*.xls*", respectcase    //获取文件夹中所有excel表的名称
dis `"`files'"'                                   //将获取的表名存入局部宏
foreach file in `files' {
	*表5
	import excel using "`file'", sheet("Service Metrics (items 3-6)") allstring clear
	gen Category="",before(A)
	replace Category=usubstr(A,1,1)  if regexm(A,"^[0-9].")&strmatch(A,"*Week*")
	replace Category="4"             if strmatch(A,"*4.*")&strmatch(A,"*Week*")
	replace Category="5"             if strmatch(A,"*5.*")&strmatch(A,"*Week*")
	replace Category="1"             if A=="1. Average Train Speed by Train Type and Overall System Average Train Speed reported in Miles per Hour (MPH)"
	gen Measure=A  if regexm(A,"(.*)(^[0-9].)")&strmatch(A,"*Week*")
	replace Measure=A  			  if strmatch(A,"*4.*")&strmatch(A,"*Week*")
	replace Measure=A             if strmatch(A,"*5.*")&strmatch(A,"*Week*")
	replace Measure=A  			  if A=="1. Average Train Speed by Train Type and Overall System Average Train Speed reported in Miles per Hour (MPH)"
	gen date="`file'"
	replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
	gen Railroad=substr("UP20141022-20140322",1,strpos("UP20141022-20140322","2014")-1) ,before(Category)
	drop in 1/4
	dropmiss, obs force  //删除整行缺失的行
	carryforward Category Measure,replace
	foreach i of numlist 5 {    //在同一个文件夹内循环
		keep if Category=="`i'"
		drop if strmatch(A,"*`i'*")&B==""
		foreach x of varlist * {
			replace `x'="Total" if `x'[_n]==""&`x'[_n-1]=="Total"
		}
		drop if (A=="Train Type"|A==" Train Type")&(B==""|B=="Cause")
		rename A Variable
		replace G=G[2]  in 1  if F[1]=="Other"&G[1]==""
		foreach x of varlist * {
			replace `x'=subinstr(`x'," ","_",.)    in 1
		}
		replace Variable=""  in 1
		replace Category=""  in 1
		replace date=""      in 1
		replace Railroad=""  in 1
		replace Measure=""   in 1   //保持这几个变量的名称不变
		nrow 1
		cap findname, all(missing(@)) 
		cap drop `r(varlist)'
		drop if Variable==""
		cap gen Other=""
		cap ren Locomotive_Power Locomotive_power
		append using $path\Clndata\表`i'
		save $path\Clndata\表`i',replace	
	}
	*表6
	import excel using "`file'", sheet("Service Metrics (items 3-6)") allstring clear
	gen Category="",before(A)
	replace Category=usubstr(A,1,1)  if regexm(A,"^[0-9].")&strmatch(A,"*Week*")
	replace Category="4"             if strmatch(A,"*4.*")&strmatch(A,"*Week*")
	replace Category="5"             if strmatch(A,"*5.*")&strmatch(A,"*Week*")
	replace Category="1"             if A=="1. Average Train Speed by Train Type and Overall System Average Train Speed reported in Miles per Hour (MPH)"
	gen Measure=A  if regexm(A,"(.*)(^[0-9].)")&strmatch(A,"*Week*")
	replace Measure=A  			  if strmatch(A,"*4.*")&strmatch(A,"*Week*")
	replace Measure=A             if strmatch(A,"*5.*")&strmatch(A,"*Week*")
	replace Measure=A  			  if A=="1. Average Train Speed by Train Type and Overall System Average Train Speed reported in Miles per Hour (MPH)"
	gen date="`file'"
	replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
	gen Railroad=substr("UP20141022-20140322",1,strpos("UP20141022-20140322","2014")-1) ,before(Category)
	drop in 1/4
	dropmiss, obs force  //删除整行缺失的行
	carryforward Category Measure,replace
	foreach i of numlist 6 {
		keep if Category=="`i'"
		*删除无关的行和列
		drop in 1/2
		replace B=B+"1"  in 1
		replace C=C+"1"  in 1  if C[1]!=""
		replace F=F+"1"  in 1  if C[1]==""
		replace D=D+"2"  in 1  if D[1]!=""
		replace G=G+"2"  in 1  if D[1]==""
		replace E=E+"2"  in 1  if E[1]!=""
		cap replace K=K+"2"  in 1  if E[1]==""
		rename A Variable
		foreach x of varlist * {
			replace `x'=subinstr(`x'," ","_",.)    in 1
		}
		replace Variable=""  in 1
		replace Category=""  in 1
		replace date=""      in 1
		replace Railroad=""  in 1
		replace Measure=""   in 1   //保持这几个变量的名称不变
		
		nrow 1
		destring Loaded1 Empty1 Loaded2 Empty2,replace force
		gen Loaded=round(Loaded1+Loaded2),before(Loaded1)
		gen Empty=round(Empty1+Empty2),before(Empty1)
		tostring Loaded* Empty*,replace force
		drop if Variable==""
		append using $path\Clndata\表`i'
		save $path\Clndata\表`i',replace
	}
	*表8
	import excel using "`file'", sheet("Grain Metrics 2 (item 8)") allstring clear
	dis `"`file'"'
	gen Category=""
	replace Category=usubstr(A,1,1) if regexm(A,"^[0-9].")
	gen Measure=""
	replace Measure=A               if regexm(A,"^[0-9].")
	gen date="`file'"
	replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
	gen Railroad=substr("UP20141022-20140322",1,strpos("UP20141022-20140322","2014")-1) ,before(Category)
	order Category Measure date,after(Railroad)
	drop in 1/4
	dropmiss ,obs force  //删除整行缺失的行
	carryforward Category Measure,replace
	foreach i of numlist 8 {
		keep if Category=="`i'"
		drop if A==""|(strmatch(A,"*`i'*")&B=="")
		rename A SubVariable
		replace SubVariable=""  in 1
		replace Category=""     in 1
		replace date=""         in 1
		replace Railroad=""     in 1
		replace Measure=""      in 1   //保持这几个变量的名称不变
		nrow 1
		cap findname, all(missing(@)) 
		cap drop `r(varlist)'
		append using $path\Clndata\表`i'
		save $path\Clndata\表`i',replace
	}
// 	*表9
// 	import excel using "`file'", sheet("Grain & Coal Plans (items 9-10)") allstring clear
// 	dis `"`file'"'
// 	gen Category=""
// 	order Category
// 	replace Category=usubstr(A,1,1) if strmatch(A,"*9.*")
// 	replace Category="10"           if strmatch(A,"*10.*")
// 	gen date="`file'"
// 	replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
// 	gen Railroad=substr("UP20141022-20140322",1,strpos("UP20141022-20140322","2014")-1) ,before(Category)
// 	dropmiss,obs force  //删除整行缺失的行
// 	carryforward Category,replace
// 	foreach i of numlist 9 {
// 		local j=`i'+1
// 		keep if Category=="`j'"
// 		*删除无关的行和列
// 		drop in 1
// 		rename A Variable
// 		drop if Variable=="10. Average Daily Coal Unit Train Loadings vs.Plan for the Reporting Week By Coal"
// 		foreach x of varlist * {
// 			replace `x'=subinstr(`x'," ","_",.)    in 1
// 		}
// 		replace Variable=""  in 1
// 		replace Category=""  in 1
// 		replace date=""      in 1
// 		replace Railroad=""  in 1
// 		nrow 1
// 		drop if Variable==""
// 		cap ren Plan Loadings_Plan
// 		cap ren Actual Loadings_Average	
// 		cap keep date Railroad Category Variable Loadings_Plan Loadings_Average			
// 		append using $path\Clndata\表`j'
// 		save $path\Clndata\表`j',replace
// 	}
}


// *1.把7个文件夹里所有文件的表1-表10分别合并起来
// local wjs "BNSF20141029-20170322 KCS20141022-20140322 CSX20141022-20170322 CP20141022-20170322 CN20141029-20170322 UP20141022-20140322"   //说明有7个文件夹
// foreach wj of local wjs {     //循环，挨个处理每个文件夹里的文件
// 	cd "$path\Sdata\transportation\\`wj'"             //指定路径
// 	local files : dir . file "*.xls*", respectcase    //获取文件夹中所有excel表的名称
// 	dis `"`files'"'                                   //将获取的表名存入局部宏
// 	foreach file in `files' {    //第二个循环，挨个处理每个文件里面的表10
// 		*表10
// 		import excel using "`file'", sheet("Grain & Coal Plans (items 9-10)") allstring clear
// 		dis `"`file'"'
// 		gen Category=""
// 		order Category
// 		replace Category=usubstr(A,1,1) if strmatch(A,"*9.*")
// 		replace Category="10"           if strmatch(A,"*10.*")
// 		gen date="`file'"
// 		replace date=ustrregexs(3) if ustrregexm(date,"(.*)(data_)(.*)(.xls.*)")
// 		gen Railroad=substr("`wj'",1,strpos("`wj'","2014")-1) ,before(Category)
// 		dropmiss,obs force  //删除整行缺失的行
// 		carryforward Category,replace
// 		foreach i of numlist 9 {
// 			keep if Category=="`i'"
// 			*删除无关的行和列
// 			drop if A=="9.      Plan vs. Performance For Grain Shuttle (Or Dedicated Grain Train) Round Trips, By Region, Updated To Reflect The Previous Four Weeks"
// 			rename A SubVariable
// 			rename B Trip_Plan
// 			rename C Trip_Performance
// 			foreach x of varlist * {
// 				replace `x'=subinstr(`x'," ","_",.)    in 1
// 			}
// 			replace SubVariable=""  in 1
// 			replace Category=""  in 1
// 			replace date=""      in 1
// 			replace Railroad=""  in 1
// 			nrow 1
// 			drop if SubVariable==""	
// 			append using $path\Clndata\表`i'
// 			save $path\Clndata\表`i',replace
// 		}
// 	}
// }

use $path\Clndata\表5 ,clear
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
save $path\Clndata\表5 ,replace