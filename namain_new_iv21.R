
library(dplyr)
library(magrittr)
library(openxlsx)
library(haven)
library(data.table)

#######################################
######### Definir parámetros ##########
#######################################

yy <- 2021 # año del trimestre
mm <- 11 # mes central

#######################################

base<-get(load(paste0("//buvmfswinp01/ene2/Automatización boletines/Bases/",yy,"-",mm,".RData")))
base<-zap_formats(base)


base%<>%mutate(rama_t=ifelse(b14_1_rev4cl_caenes==1, 1, 0),
               rama_t=ifelse(b14_1_rev4cl_caenes %in% 2:5, 2, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes==3, 3, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes==6, 4, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes %in% 7:9, 5, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes==10, 6, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes==11, 7, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes==12, 8, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes %in% 13:14, 9, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes %in% 15:17, 10, rama_t),
               rama_t=ifelse(b14_1_rev4cl_caenes %in% 18:21, 11, rama_t))

####################### Personas ################################

# total
t1<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 1]))
t2<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 2:5]))
t3<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 3]))
t4<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 6]))
t5<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 7:9]))
t6<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 10]))
t7<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 11]))
t8<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 12]))
t9<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 13:14]))
t10<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 15:17]))
t11<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 18:21]))
tt<-base%>%summarise(rama_t2=sum(fact_cal[b14_1_rev4cl_caenes %in% 1:21]))

e_vac<-""

tbl_q_t2<-cbind(tt, e_vac, e_vac, t1, e_vac, e_vac, t2, e_vac, e_vac, t3, e_vac, e_vac, t4, e_vac, e_vac, t5, e_vac, e_vac, t6, e_vac, e_vac, t7, e_vac, e_vac, 
                t8, e_vac, e_vac, t9, e_vac, e_vac, t10, e_vac, e_vac, t11, e_vac, e_vac)



# employers
t1<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 1]))
t2<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 2:5]))
t3<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 3]))
t4<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 6]))
t5<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 7:9]))
t6<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 10]))
t7<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 11]))
t8<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 12]))
t9<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 13:14]))
t10<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 15:17]))
t11<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 18:21]))
tt<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & b14_1_rev4cl_caenes %in% 1:21]))

tbl_s_t2<-cbind(tt, e_vac, e_vac, t1, e_vac, e_vac, t2, e_vac, e_vac, t3, e_vac, e_vac, t4, e_vac, e_vac, t5, e_vac, e_vac, t6, e_vac, e_vac, t7, e_vac, e_vac, 
                t8, e_vac, e_vac, t9, e_vac, e_vac, t10, e_vac, e_vac, t11, e_vac, e_vac)


# employees
t1<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 1]))
t2<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 2:5]))
t3<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 3]))
t4<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 6]))
t5<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 7:9]))
t6<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 10]))
t7<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 11]))
t8<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 12]))
t9<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 13:14]))
t10<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 15:17]))
t11<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 18:21]))
tt<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & b14_1_rev4cl_caenes %in% 1:21]))

tbl_e_t2<-cbind(tt, e_vac, e_vac, t1, e_vac, e_vac, t2, e_vac, e_vac, t3, e_vac, e_vac, t4, e_vac, e_vac, t5, e_vac, e_vac, t6, e_vac, e_vac, t7, e_vac, e_vac, 
                t8, e_vac, e_vac, t9, e_vac, e_vac, t10, e_vac, e_vac, t11, e_vac, e_vac)


tbl_personas<-cbind(tbl_q_t2, tbl_e_t2, tbl_s_t2)


write.xlsx(tbl_personas, "//buvmfswinp01/ene2/Repo/T0111_NP.xlsx", row.names = F, overwrite = TRUE)
write.xlsx(tbl_personas, "//buvmfswinp01/ene2/Repo/T0111_NP.xlsx", row.names = F, overwrite = TRUE)

# copiar toda la fila de este archivo "T0111_NP.xlsx" y pegarla en archivo NAMAIN_T0111, hoja "_NP", según trimestre 


#################################################################
#################################################################
####################### Horas tr ################################
#################################################################
#################################################################

# total
t1<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 1]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 1]))
t2<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 2:5]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 2:5]))
t3<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 3]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 3]))
t4<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 6]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 6]))
t5<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 7:9]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 7:9]))
t6<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 10]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 10]))
t7<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 11]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 11]))
t8<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 12]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 12]))
t9<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 13:14]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 13:14]))
t10<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 15:17]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 15:17]))
t11<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 18:21]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 18:21]))
tt<-base%>%summarise(rama_t2=sum(fact_cal[habituales <= 168 & b14_1_rev4cl_caenes %in% 1:21]*habituales[habituales <= 168 & b14_1_rev4cl_caenes %in% 1:21]))

e_vac<-""

tbl_q_t2<-cbind(tt, e_vac, e_vac, t1, e_vac, e_vac, t2, e_vac, e_vac, t3, e_vac, e_vac, t4, e_vac, e_vac, t5, e_vac, e_vac, t6, e_vac, e_vac, t7, e_vac, e_vac, 
                t8, e_vac, e_vac, t9, e_vac, e_vac, t10, e_vac, e_vac, t11, e_vac, e_vac)



# employers
t1<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 1]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 1]))
t2<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 2:5]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 2:5]))
t3<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 3]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 3]))
t4<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 6]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 6]))
t5<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 7:9]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 7:9]))
t6<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 10]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 10]))
t7<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 11]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 11]))
t8<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 12]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 12]))
t9<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 13:14]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 13:14]))
t10<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 15:17]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 15:17]))
t11<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 18:21]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 18:21]))
tt<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 1:21]*habituales[cise %in% 1:2 & habituales <= 168 & b14_1_rev4cl_caenes %in% 1:21]))

tbl_s_t2<-cbind(tt, e_vac, e_vac, t1, e_vac, e_vac, t2, e_vac, e_vac, t3, e_vac, e_vac, t4, e_vac, e_vac, t5, e_vac, e_vac, t6, e_vac, e_vac, t7, e_vac, e_vac, 
                t8, e_vac, e_vac, t9, e_vac, e_vac, t10, e_vac, e_vac, t11, e_vac, e_vac)


# employees
t1<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 1]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 1]))
t2<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 2:5]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 2:5]))
t3<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 3]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 3]))
t4<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 6]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 6]))
t5<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 7:9]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 7:9]))
t6<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 10]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 10]))
t7<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 11]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 11]))
t8<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 12]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 12]))
t9<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 13:14]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 13:14]))
t10<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 15:17]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 15:17]))
t11<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 18:21]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 18:21]))
tt<-base%>%summarise(rama_t2=sum(fact_cal[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 1:21]*habituales[cise %in% 3:7 & habituales <= 168 & b14_1_rev4cl_caenes %in% 1:21]))

tbl_e_t2<-cbind(tt, e_vac, e_vac, t1, e_vac, e_vac, t2, e_vac, e_vac, t3, e_vac, e_vac, t4, e_vac, e_vac, t5, e_vac, e_vac, t6, e_vac, e_vac, t7, e_vac, e_vac, 
                t8, e_vac, e_vac, t9, e_vac, e_vac, t10, e_vac, e_vac, t11, e_vac, e_vac)


tbl_horas<-cbind(tbl_q_t2, tbl_e_t2, tbl_s_t2)


write.xlsx(tbl_horas, "//buvmfswinp01/ene2/Repo/T0111_NH.xlsx", row.names = F, overwrite = TRUE)
# copiar toda la fila de este archivo "T0111_NH.xlsx" y pegarla en archivo NAMAIN_T0111, hoja "_NH", según trimestre 
