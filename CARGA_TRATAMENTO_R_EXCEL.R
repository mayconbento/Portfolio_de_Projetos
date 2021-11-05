# Minha versão do RStudio
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

source('limpar_dados_final_x1.R'  ,encoding = 'UTF8')
source('limpar_dados_final_x2.R'  ,encoding = 'UTF8')
source('limpar_dados_final_x3.R'  ,encoding = 'UTF8')
source('limpar_dados_final_x4.R'  ,encoding = 'UTF8')
source('limpar_dados_final_x5.R'  ,encoding = 'UTF8')
source('limpar_dados_final_x6.R'  ,encoding = 'UTF8')
source('limpar_dados_final_x7.R'  ,encoding = 'UTF8')

# unir arquivos xlsx
rm(list = ls(all.names = TRUE))

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

setwd(paste0(letra,':/R_STUDIO/resultado_por_empresa'))

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

consolidado_dia<- aggregate(`R$ TOTAL` ~ DIA+ empresa+CL, 
                            data = consolidado_dia, 
                            FUN = sum, na.rm = TRUE)  

str(consolidado)

consolidado$DATA<- as.Date(consolidado$DATA , origin= '1899-12-30')

consolidado<- consolidado[consolidado$CL !='01',]


openxlsx::write.xlsx(consolidado,
                     paste0(letra,':/R_STUDIO/consolidado/consolidado_mes.xlsx'))

#rm(list = ls(all.names = TRUE))

#********************* unir INFO2 *************************
#*******************************************************
# unir arquivos XLSX

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

setwd(paste0(letra,':/R_STUDIO/INFO2_por_empresa'))

library('readxl')
library('dplyr')

file.list <- list.files(pattern='*.xlsx')
file.list

consolidado_INFO2 <- bind_rows(lapply(file.list, read_excel))

linhas<- dim(consolidado_INFO2)
linhas[1]

sum(consolidado_INFO2$TOTAL_INFO2)

getwd()

openxlsx::write.xlsx(consolidado_INFO2, 'D:/R_STUDIO/INFO2_por_empresa/consolidado/consolidado_INFO2.xlsx')
           
consolidado_INFO2<- subset(consolidado_INFO2, select = c(empresa, TOTAL_INFO2, FROTA))

sum(consolidado_INFO2$TOTAL_INFO2)
data.frame(consolidado_INFO2)

consolidado_INFO2<- aggregate(TOTAL_INFO2 ~ empresa, data = consolidado_INFO2, 
                           FUN = sum, na.rm = TRUE)

data.frame(consolidado_INFO2)

#agora<- Sys.time()

#consolidado_INFO2$Data<- agora

data.frame(consolidado_INFO2)

consolidado_INFO2[is.na(consolidado_INFO2)]<- 0

#openxlsx::write.xlsx(consolidado_INFO2,paste0(letra,':/R_STUDIO/consolidado_INFO2/consolidado_INFO2.xlsx'))

#********************* unir FECHAM_RESULT *************************
#*******************************************************
# unir arquivos XLSX

getwd()

letra<- getwd()

letra<- str_sub(letra, start =1, end = 1) 
letra

setwd(paste0(letra,':/R_STUDIO/FECHAM_RESULT_por_empresa'))

library('readxl')
library('dplyr')

file.list <- list.files(pattern='*.xlsx')
file.list

FECHAM_RESULT_consolidado <- bind_rows(lapply(file.list, read_excel))

linhas<- dim(FECHAM_RESULT_consolidado)
linhas[1]

#********************** EXTRAIR CUSTO 01,02,03************
#********************************************************************

FECHAM_RESULT_consolidado<- FECHAM_RESULT_consolidado[,c(2,1,3,4,5)]

CUSTO<- c('01','02','03')

CUSTO_INFO2<- FECHAM_RESULT_consolidado[FECHAM_RESULT_consolidado$CL %in% CUSTO,]
CUSTO_INFO2<- CUSTO_INFO2[,c(1,3,5)]

CUSTO_INFO2<- aggregate(`R$ TOTAL`~ empresa, data = CUSTO_INFO2, 
                           FUN = sum, na.rm = TRUE)

CUSTO_INFO2<- left_join(consolidado_INFO2, CUSTO_INFO2)
head(CUSTO_INFO2,10)

CUSTO_INFO2<- CUSTO_INFO2[complete.cases(CUSTO_INFO2),]


CUSTO_INFO2$PERCENTUAL<- round((CUSTO_INFO2$`R$ TOTAL`/CUSTO_INFO2$TOTAL_INFO2),3)
CUSTO_INFO2

nomes<- colnames(CUSTO_INFO2)

##************custo geral****************
CUSTO <- sum(CUSTO_INFO2$`R$ TOTAL`)
INFO2_geral <- sum(CUSTO_INFO2$TOTAL_INFO2)
geral<- round(sum(CUSTO_INFO2$`R$ TOTAL`)/(sum(CUSTO_INFO2$TOTAL_INFO2)),3)

row<- data.frame(empresa=c('TOTAL'),
               TOTAL_INFO2=c(INFO2_geral),
               `R$ TOTAL`=c(CUSTO),
               PERCENTUAL=c(geral))

colnames(row)<- nomes

CUSTO_INFO2<- bind_rows(CUSTO_INFO2, row)
CUSTO_INFO2

getwd()
openxlsx::write.xlsx(CUSTO_INFO2,paste0(letra,':/R_STUDIO/consolidado_INFO2/consolidado_CUSTO_INFO2.xlsx'))

FECHAM_RESULT_consolidado<- pivot_wider(data = FECHAM_RESULT_consolidado, 
                              names_from = `empresa`,
                              values_from = `R$ TOTAL`)

head(FECHAM_RESULT_consolidado)

FECHAM_RESULT_consolidado[is.na(FECHAM_RESULT_consolidado)]<- 0

FECHAM_RESULT_consolidado$TOTAL<- apply(FECHAM_RESULT_consolidado[,c(4:10)],1, sum)

FECHAM_RESULT_consolidado$CUSTEIO<- NULL

FECHAM_RESULT_consolidado<- FECHAM_RESULT_consolidado[order(FECHAM_RESULT_consolidado$CL),]

FECHAM_RESULT_consolidado<- FECHAM_RESULT_consolidado[FECHAM_RESULT_consolidado$CL !='91',]

#*****************************Setting Styles**********************************************
p <- createWorkbook()

# Add some sheets to the workbook
addWorksheet(p, "consolidado_mes", zoom = 95)
addWorksheet(p, "FECHAM_RESULT_consolidado")
addWorksheet(p, "custo_INFO2")

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

setRowHeights(p, sheet = "FECHAM_RESULT_consolidado", heights = 18, rows = 1:17)

width_vec <- apply(FECHAM_RESULT_consolidado, 2, function(x) max(nchar(as.character(x)) + 3, na.rm = TRUE)) 

setColWidths(p, sheet = "FECHAM_RESULT_consolidado", 
             cols = 1:ncol(FECHAM_RESULT_consolidado), widths = width_vec)

#setColWidths(p, sheet = "FECHAM_RESULT_consolidado", 
#             cols = c(3,4,5,6,7,8,9,10), widths = c(12.57,12.57,12.57,12.57,12.57,12.57,12.57,14),
#             hidden = FALSE, ignoreMergedCells = TRUE)

setColWidths(p, sheet = "FECHAM_RESULT_consolidado", 
             cols = c(2), widths = c(25.43),
             hidden = FALSE, ignoreMergedCells = TRUE)

setColWidths(p, sheet = "FECHAM_RESULT_consolidado", 
             cols = c(1), widths = c(4.3),
             hidden = FALSE, ignoreMergedCells = TRUE)

addStyle(wb = p, sheet = "FECHAM_RESULT_consolidado" , style = headerstyle, rows = 1:1,  cols = 1:10, gridExpand = TRUE)
addStyle(wb = p, sheet = "FECHAM_RESULT_consolidado" , style = bodystyle,   rows = 2:17, cols = 1:10, gridExpand = TRUE)
addStyle(wb = p, sheet = "FECHAM_RESULT_consolidado" , style = cellstyle,   rows = 2:17, cols = 1:2,  gridExpand = TRUE)

freezePane(wb = p, sheet = "consolidado_mes", firstActiveRow = 2)
freezePane(wb = p, sheet = "FECHAM_RESULT_consolidado", firstActiveRow = 2)

writeData(p, sheet = "consolidado_mes", consolidado)
writeData(p, sheet = "FECHAM_RESULT_consolidado", FECHAM_RESULT_consolidado)

agora<- str_replace(Sys.time(), pattern = " ", " ->> ")
agora<- as.character(agora)

writeData(p, sheet = "FECHAM_RESULT_consolidado", xy = c(11,1), agora)
writeData(p, sheet = "custo_INFO2", CUSTO_INFO2)
getwd()

setwd(paste0(letra,':/R_STUDIO/FECHAM_RESULT_consolidado'))

saveWorkbook(p, "FECHAM_RESULT_consolidado.xlsx", overwrite = TRUE)
openXL("FECHAM_RESULT_consolidado.xlsx")

dupli<- consolidado[duplicated(consolidado),]
dim(unique(consolidado))
dim(distinct(consolidado))

