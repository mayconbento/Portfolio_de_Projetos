# R version 3.6.2 (2019-12-12) -- "Dark and Stormy Night"

# Bilbliotecas
library('stringr')
library('readxl')
library('dplyr')
library('magrittr')
library('dplyr')
library('tidyr')
library('openxlsx')

# LIMPAR VARIÁVEIS
rm(list = ls(all.names = TRUE))

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

#options(OutDec = ".", )

dir<- paste0(letra,':/R_STUDIO/unir_arquivos')
setwd(dir)
getwd()

source('limpar_dados_X1.R'  ,encoding = 'UTF8')
source('limpar_dados_X2.R'  ,encoding = 'UTF8')
source('limpar_dados_X3.R'  ,encoding = 'UTF8')
source('limpar_dados_X4.R'  ,encoding = 'UTF8')
source('limpar_dados_X5.R'  ,encoding = 'UTF8')
source('limpar_dados_X6.R'  ,encoding = 'UTF8')
source('limpar_dados_X7.R'  ,encoding = 'UTF8')

# unir arquivos xlsx
rm(list = ls(all.names = TRUE))

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

setwd(paste0(letra,':/R_STUDIO/resultado_por_drive'))

library('readxl')
library('dplyr')

file.list <- list.files(pattern='*.xlsx')
file.list

infos<- data.frame(file.info(file.list))

INFO<- separate(data = infos, 
                col = mtime, 
                into = c('Data',
                         'Modificado em:'),sep = " ")

INFO %>% select(c(4,5))

consolidado <- bind_rows(lapply(file.list, read_excel))

linhas<- dim(consolidado)
linhas[1]

dim(consolidado)-dim(unique(consolidado))

consolidado_dia<- consolidado
consolidado_dia$DIA<- lubridate::day(consolidado_dia$DATA)

consolidado_dia<- consolidado_dia[  consolidado_dia$CL=='01'|
                                    consolidado_dia$CL=='02'|
                                    consolidado_dia$CL=='03',]

consolidado_dia<- aggregate(`R$ TOTAL` ~ DIA+ DRIVE+CL, data = consolidado_dia, 
                             FUN = sum, na.rm = TRUE)  

str(consolidado)

consolidado$DATA<- as.Date(consolidado$DATA , origin= '1899-12-30')

consolidado<- consolidado[consolidado$CL !='91',]

openxlsx::write.xlsx(consolidado,paste0(letra,':/R_STUDIO/consolidado/consolidado_mes.xlsx'))

#rm(list = ls(all.names = TRUE))

#********************* unir KM *************************
#*******************************************************
# unir arquivos XLSX

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

setwd(paste0(letra,':/R_STUDIO/km_por_drive'))

library('readxl')
library('dplyr')

file.list <- list.files(pattern='*.xlsx')
file.list

consolidado_km <- bind_rows(lapply(file.list, read_excel))

linhas<- dim(consolidado_km)
linhas[1]

sum(consolidado_km$TOTAL_KM)

getwd()

openxlsx::write.xlsx(consolidado_km, 'D:/R_STUDIO/km_por_drive/consolidado/consolidado_km.xlsx')
           
consolidado_km<- subset(consolidado_km, select = c(DRIVE, TOTAL_KM, FROTA))

sum(consolidado_km$TOTAL_KM)
data.frame(consolidado_km)

consolidado_km<- aggregate(TOTAL_KM ~ DRIVE, data = consolidado_km, 
                           FUN = sum, na.rm = TRUE)

data.frame(consolidado_km)

#agora<- Sys.time()

#consolidado_km$Data<- agora

data.frame(consolidado_km)

consolidado_km[is.na(consolidado_km)]<- 0

#openxlsx::write.xlsx(consolidado_km,paste0(letra,':/R_STUDIO/consolidado_km/consolidado_KM.xlsx'))

#********************* unir DRE *************************
#*******************************************************
# unir arquivos XLSX

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

setwd(paste0(letra,':/R_STUDIO/DRO_por_drive'))

library('readxl')
library('dplyr')

file.list <- list.files(pattern='*.xlsx')
file.list

DRE_consolidado <- bind_rows(lapply(file.list, read_excel))

linhas<- dim(DRE_consolidado)
linhas[1]

#********************** EXTRAIR CUSTO MANUTENÇÃO 01,02,03************
#********************************************************************

DRO_consolidado<- DRO_consolidado[,c(2,1,3,4,5)]

custo_manut<- c('01','02','03')

CUSTO_KM<- DRO_consolidado[DRO_consolidado$CL %in% custo_manut,]
CUSTO_KM<- CUSTO_KM[,c(1,3,5)]

CUSTO_KM<- aggregate(`R$ TOTAL`~ DRIVE, data = CUSTO_KM, 
                           FUN = sum, na.rm = TRUE)

CUSTO_KM<- left_join(consolidado_km, CUSTO_KM)
head(CUSTO_KM,10)

CUSTO_KM<- CUSTO_KM[complete.cases(CUSTO_KM),]

CUSTO_KM$PERCENTUAL<- round((CUSTO_KM$`R$ TOTAL`/CUSTO_KM$TOTAL_KM),3)
CUSTO_KM
nomes<- colnames(CUSTO_KM)

##************custo geral****************
custo_manut <- sum(CUSTO_KM$`R$ TOTAL`)
km_geral <- sum(CUSTO_KM$TOTAL_KM)
geral<- round(sum(CUSTO_KM$`R$ TOTAL`)/(sum(CUSTO_KM$TOTAL_KM)),3)

row<- data.frame(DRIVE=c('TOTAL'),
               TOTAL_KM=c(km_geral),
               `R$ TOTAL`=c(custo_manut),
               PERCENTUAL=c(geral))

colnames(row)<- nomes

CUSTO_KM<- bind_rows(CUSTO_KM, row)
CUSTO_KM

getwd()
openxlsx::write.xlsx(CUSTO_KM,paste0(letra,':/R_STUDIO/consolidado_km/consolidado_CUSTO_KM.xlsx'))

DRO_consolidado<- pivot_wider(data = DRO_consolidado, 
                              names_from = `DRIVE`,
                              values_from = `R$ TOTAL`)

head(DRO_consolidado)

DRO_consolidado[is.na(DRO_consolidado)]<- 0

DRO_consolidado$TOTAL<- apply(DRO_consolidado[,c(4:10)],1, sum)

DRO_consolidado$CUSTEIO<- NULL

DRO_consolidado<- DRO_consolidado[order(DRO_consolidado$CL),]

DRO_consolidado<- DRO_consolidado[DRO_consolidado$CL !='91',]

#*****************************Setting Styles**********************************************
#**********************inserindo informações no excel e formatando worksheets*************
p <- createWorkbook()

# Add some sheets to the workbook
addWorksheet(p, "consolidado_mes", zoom = 95)
addWorksheet(p, "dro_consolidado")
addWorksheet(p, "custo_km")

headerstyle<- createStyle(
  fontSize = 11, fontName = "Calibri",textDecoration = "bold",
  halign = "center", fgFill = 'lightblue2', 
  border = "TopBottomLeftRight", borderColour = "black", borderStyle = "thin")

bodystyle <- createStyle(
  fontSize = 11, fontName = "Calibri", numFmt = "0,0.00",
  halign = "right", fgFill = 'white', 
  border = "TopBottomLeftRight", borderColour = "black", borderStyle = "hair")

cellstyle<- createStyle(
  fontSize = 11, fontName = "Calibri",
  halign = "left", fgFill = 'white', 
  border = "TopBottomLeftRight", borderColour = "black", borderStyle = "hair")

setRowHeights(p, sheet = "dro_consolidado", heights = 18, rows = 1:17)

width_vec <- apply(DRO_consolidado, 2, function(x) max(nchar(as.character(x)) + 3, na.rm = TRUE)) 

setColWidths(p, sheet = "dro_consolidado", 
             cols = 1:ncol(DRO_consolidado), widths = width_vec)

#setColWidths(p, sheet = "dro_consolidado", 
#             cols = c(3,4,5,6,7,8,9,10), widths = c(12.57,12.57,12.57,12.57,12.57,12.57,12.57,14),
#             hidden = FALSE, ignoreMergedCells = TRUE)

setColWidths(p, sheet = "dro_consolidado", 
             cols = c(2), widths = c(25.43),
             hidden = FALSE, ignoreMergedCells = TRUE)

setColWidths(p, sheet = "dro_consolidado", 
             cols = c(1), widths = c(4.3),
             hidden = FALSE, ignoreMergedCells = TRUE)

addStyle(wb = p, sheet = "dro_consolidado" , style = headerstyle, rows = 1:1,  cols = 1:10, gridExpand = TRUE)
addStyle(wb = p, sheet = "dro_consolidado" , style = bodystyle,   rows = 2:17, cols = 1:10, gridExpand = TRUE)
addStyle(wb = p, sheet = "dro_consolidado" , style = cellstyle,   rows = 2:17, cols = 1:2,  gridExpand = TRUE)

freezePane(wb = p, sheet = "consolidado_mes", firstActiveRow = 2)
freezePane(wb = p, sheet = "dro_consolidado", firstActiveRow = 2)

writeData(p, sheet = "consolidado_mes", consolidado)
writeData(p, sheet = "dro_consolidado", DRO_consolidado)

agora<- str_replace(Sys.time(), pattern = " ", " ->> ")
agora<- as.character(agora)

writeData(p, sheet = "dro_consolidado", xy = c(11,1), agora)
writeData(p, sheet = "custo_km", CUSTO_KM)
getwd()

setwd(paste0(letra,':/R_STUDIO/DRO_consolidado'))

saveWorkbook(p, "DRO_consolidado.xlsx", overwrite = TRUE)
openXL("DRO_consolidado.xlsx")

dupli<- consolidado[duplicated(consolidado),]
dim(unique(consolidado))
dim(distinct(consolidado))
