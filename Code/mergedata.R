#1.准备工作：指定工作目录，调用所需R包等
install.packages("tidyverse")
setwd("C:\\Users\\li_zhi\\Desktop\\From\\Shan")
library(xlsx)
library(tidyverse)
#2.导入数据
read.csv("Data\\Sdata\\New folder\\1-Train_Speeds.csv")%>%t()
#3.处理数据成标准格式
#4.合并各个数据表
