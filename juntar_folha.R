library('readxl')
library('ggplot2')
library('openxlsx')
library('magrittr')
library('dplyr')
library('stringr')
library('openxlsx')
library('tidyr')
#library('tidyverse')
library('purrr')

options(digits = 5)

# LIMPAR VARIÁVEIS
rm(list = ls(all.names = TRUE))

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

dir<- paste0(letra,':/R_STUDIO/Folha')

setwd(dir)

getwd()

source('folha_CAP1.R',   encoding = 'UTF8')
source('folha_CAP7.R',   encoding = 'UTF8')
source('folha_LR13.R',   encoding = 'UTF8')
source('folha_RLC16.R',  encoding = 'UTF8')
source('folha_WS4.R',    encoding = 'UTF8')
source('folha_RS16.R',   encoding = 'UTF8')

# LIMPAR VARIÁVEIS
rm(list = ls(all.names = TRUE))

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

dir<- paste0(letra,':/R_STUDIO/Folha/DRO_FOLHA_DRIVE')

setwd(dir)

# unir arquivos excel na pasta
file.list <- list.files(pattern='*.xlsx')
file.list

DRO_FOLHA_CONSOLIDADO <- bind_rows(lapply(file.list, read_excel))

str(DRO_FOLHA_CONSOLIDADO)

DRO_FOLHA_CONSOLIDADO<- DRO_FOLHA_CONSOLIDADO[,c(5,1,2,3,4)]

#write.xlsx(DRO_FOLHA_CONSOLIDADO, 'DRO_FOLHA_CONSOL.xlsx')

#unique(DRO_FOLHA_CONSOLIDADO$DRIVE)

DRO_FOLHA_CONSOLIDADO[is.na(DRO_FOLHA_CONSOLIDADO)]<- 0


DRO_WIDER<- pivot_wider(data = DRO_FOLHA_CONSOLIDADO, 
                   names_from = DRIVE,
                   values_from = VALOR)

DRO_WIDER[is.na(DRO_WIDER)]<- 0

DRO_WIDER$TOTAL<- apply(DRO_WIDER[,c(4:9)],1, sum)

ordem<- read_xlsx('D:/R_STUDIO/Folha/ordem.xlsx')

DRO_WIDER<- left_join(DRO_WIDER, ordem)

DRO_WIDER<- DRO_WIDER[order(DRO_WIDER$order),]
DRO_WIDER$order<- NULL

#DRO_WIDER<- DRO_WIDER[order(DRO_WIDER$order),]
DRO_WIDER<- DRO_WIDER[, c(1,2,3,9,4,8,7,5,6,10)]

openxlsx::write.xlsx(DRO_WIDER,
                     paste0(letra,
                            ':/R_STUDIO/Folha/DRO_FOLHA_CONSOLIDADO/DRO_FOLHA_CONSOLIDADO_AGO-2021.xlsx'))

