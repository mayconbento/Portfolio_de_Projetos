
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

dir<- paste0(letra,':/R_STUDIO/Folha/DRO_FOLHA_BRUTO')

setwd(dir)

getwd()

folha<- read_xlsx('RELATORIO AGO21.xlsx', col_names = FALSE, 
                  sheet = 'X1')


NOME_DRIVE<- folha[1,1]
NOME_DRIVE

folha<- folha[-1,]

colnames(folha)<- folha[1,]

folha<- folha[-1,]

folha<- folha[,c(1:9)]

folha$TIPO[is.na(folha$EVENTO)]<- 'x'

folha$TIPO[folha$DESCRICAO=='TOTALIZACAO EMPRESA']<- 'TOTALIZACAO EMPRESA'

folha$TIPO[folha$DESCRICAO=='01 - ADMINISTRACAO']<- 'ADMINISTRACAO'

folha$TIPO[folha$DESCRICAO=='02 - MANUTENCAO']<- 'MANUTENCAO'

#folha$TIPO[folha$DESCRICAO=='03 - TRAFEGO']<- 'TRAFEGO'

colnames(folha)<- c("EVENTO","DESCRICAO","SINAL","HORAS",
                    "ATIVOS","DESLIGADOS","TOTAL",
                    "F_ATIVOS","F_FERIAS","TIPO")   

unique(folha$TIPO)

padrao<- "^[0-3]{2}\\s[0-9]{2}\\s[0-9]{2}\\s-\\s[A-Z]+"

folha<- folha %>% 
  mutate(TIPO = ifelse(str_detect(folha$DESCRICAO, pattern = padrao),
                       DESCRICAO, TIPO))

unique(folha$TIPO)

folha$TIPO[str_ends(string = folha$TIPO,pattern = ' - ADMINISTRACAO| TRAFEGO')]<- 'OUTROS'

unique(folha$TIPO)

folha<- folha %>% fill(TIPO)

folha$TIPO<-str_squish(gsub(pattern = '[^A-Z]+'  , 
                                    replacement =  ' ', 
                                    x = folha$TIPO))

unique(folha$TIPO)

folha$TIPO[str_starts(string = folha$TIPO,pattern = 'AGENC|FISCAL|OUTROS')]<- 'OUTROS'
unique(folha$TIPO)

folha$TIPO[str_starts(string = folha$TIPO,pattern = 'MOTOR')]<- 'MOTORISTAS'
unique(folha$TIPO)

folha$TIPO[str_ends(string = folha$TIPO,pattern = ' TRAFEGO|OPERACAO ESPECIAL')]<- 'OUTROS'
unique(folha$TIPO)

folha<- folha[folha$TIPO %in% 
                c('TOTALIZACAO EMPRESA','ADMINISTRACAO','MANUTENCAO',
                  #'TRAFEGO', 
                  'MOTORISTAS', 'COBRADORES','OUTROS'),]
folha

colnames(folha)

str(folha)

folha2<- folha

folha<- folha[,c(-8,-9)]
folha<- folha[complete.cases(folha),]

folha<- folha %>% mutate_at(c("HORAS","ATIVOS","DESLIGADOS","TOTAL"),
                            as.numeric)
#View(folha)
str(folha)

#folha$EVENTO<- as.factor(folha$EVENTO)
#folha$TIPO<- as.factor(folha$TIPO)

EVENT_TOTAL<- c('300','3','16','107','7','8','9','15','19','32','34','37',
         '39','33','47','160','164')

ATIVOS  <- folha[folha$EVENTO %in% EVENT_TOTAL,]
dim(ATIVOS)

ATIVOS<- aggregate(ATIVOS ~ EVENTO+DESCRICAO+TIPO, data = ATIVOS, 
                   FUN = sum, na.rm = TRUE)
dim(ATIVOS)

`%not_in%` <- purrr::negate(`%in%`)
DEDUCAO_ATIVOS<- ATIVOS[ATIVOS$EVENTO %not_in% '300',]
colnames(DEDUCAO_ATIVOS)<- c('EVENTO','DESCRICAO','TIPO','DEDUCAO')


DEDUCAO_ATIVOS
DEDUCAO_ATIVOS<- aggregate(DEDUCAO ~ TIPO, data = DEDUCAO_ATIVOS, 
                   FUN = sum, na.rm = TRUE)

DEDUCAO_ATIVOS

EVENT_PROVENTOS<- c('300')
CUSTO_PESSOAL<- ATIVOS[ATIVOS$EVENTO %in% EVENT_PROVENTOS,]
CUSTO_PESSOAL<- left_join(CUSTO_PESSOAL,DEDUCAO_ATIVOS)
CUSTO_PESSOAL

############################
# inserir "0" para NA
CUSTO_PESSOAL$DEDUCAO[is.na(CUSTO_PESSOAL$DEDUCAO)]<- 0
CUSTO_PESSOAL
############################

CUSTO_PESSOAL<- CUSTO_PESSOAL[,-1]

CUSTO_PESSOAL<- transform(CUSTO_PESSOAL,  VALOR = ATIVOS - DEDUCAO)
CUSTO_PESSOAL$DEDUCAO<- NULL
CUSTO_PESSOAL$ATIVOS<- NULL

CUSTO_PESSOAL<- CUSTO_PESSOAL[order(CUSTO_PESSOAL$VALOR),]
CUSTO_PESSOAL<- CUSTO_PESSOAL[,c(2,1,3)]
CUSTO_PESSOAL$DESCRICAO<- rep('CUSTO_COM_PESSOAL')
CUSTO_PESSOAL
#***********************
# HORAS EXTRAS FUNCIONARIOS ATIVOS E DESLIGADOS

EVENT_H.E_DESL<- c('3','16','47', '107')

HORAS_EXTRAS_DESLIGADOS  <- folha[folha$EVENTO %in% EVENT_H.E_DESL,]

HORAS_EXTRAS_DESLIGADOS

HORAS_EXTRAS_DESLIGADOS<- aggregate(DESLIGADOS ~ EVENTO+DESCRICAO+TIPO, 
                                    data = HORAS_EXTRAS_DESLIGADOS,
                                    FUN = sum, na.rm = TRUE)

HORAS_EXTRAS_DESLIGADOS$DESCRICAO<- rep('HORAS_EXTRA')

HORAS_EXTRAS_DESLIGADOS<- HORAS_EXTRAS_DESLIGADOS[,-1]

HORAS_EXTRAS_DESLIGADOS<- aggregate(DESLIGADOS ~ TIPO+DESCRICAO, data = HORAS_EXTRAS_DESLIGADOS, 
                                    FUN = sum, na.rm = TRUE)

HORAS_EXTRAS_DESLIGADOS

EVENT_H.E_ATIVOS<- c('3','16','47','107')

HORAS_EXTRAS_ATIVOS  <- folha[folha$EVENTO %in% EVENT_H.E_ATIVOS,]

#HORAS_EXTRAS_ATIVOS$DESCRICAO<- rep('HORA_EXTRA')

HORAS_EXTRAS_ATIVOS<- aggregate(ATIVOS ~ EVENTO+DESCRICAO+TIPO,
                                data = HORAS_EXTRAS_ATIVOS,
                                FUN = sum, na.rm = TRUE)

HORAS_EXTRAS_ATIVOS

HORAS_EXTRAS_ATIVOS<- HORAS_EXTRAS_ATIVOS[,-1:-2]

HORAS_EXTRAS_ATIVOS<- aggregate(ATIVOS ~ TIPO, data = HORAS_EXTRAS_ATIVOS, 
                                FUN = sum, na.rm = TRUE)

HORAS_EXTRAS_ATIVOS

HORAS_EXTRAS<- left_join(HORAS_EXTRAS_ATIVOS, HORAS_EXTRAS_DESLIGADOS)
HORAS_EXTRAS
HORAS_EXTRAS<- mutate(.data = HORAS_EXTRAS, VALOR = ATIVOS+DESLIGADOS)
HORAS_EXTRAS<- HORAS_EXTRAS[,-2]
HORAS_EXTRAS<- HORAS_EXTRAS[,-3]

HORAS_EXTRAS<- HORAS_EXTRAS[order(HORAS_EXTRAS$VALOR),]
HORAS_EXTRAS

PADRAO<- data.frame(TIPO= c('ADMINISTRACAO', 
                            'MANUTENCAO',
                            #'TRAFEGO',
                            'MOTORISTAS',
                            'COBRADORES',
                            'OUTROS',
                            'TOTALIZACAO EMPRESA'))

PADRAO

HORAS_EXTRAS<- full_join(HORAS_EXTRAS, PADRAO)
HORAS_EXTRAS
HORAS_EXTRAS<- HORAS_EXTRAS %>% fill(DESCRICAO)
HORAS_EXTRAS
HORAS_EXTRAS$VALOR[is.na(HORAS_EXTRAS$VALOR)]<- 0
HORAS_EXTRAS

#*****************TOTAL PROVENTOS*****************
TOTAL_PROV<- c('300')

TOTAL_PROVENTOS <- folha[folha$EVENTO %in% TOTAL_PROV &
                           folha$TIPO=='TOTALIZACAO EMPRESA',]

(TOTAL_PROVENTOS)

TOTAL_PROVENTOS<- TOTAL_PROVENTOS[,c(8,2,7)]

colnames(TOTAL_PROVENTOS)<- c('TIPO','DESCRICAO','VALOR')

TOTAL_PROVENTOS

#****************************************************
#****************************************************
# NÚMERO DE FUNCIONARIOS

N_FUNC <-folha2[is.na(folha2$EVENTO),]

N_FUNC<- N_FUNC[,8:10]

N_FUNC$F_ATIVOS[is.na(N_FUNC$F_ATIVOS)]<- 0
N_FUNC$F_FERIAS[is.na(N_FUNC$F_FERIAS)]<- 0
N_FUNC

N_FUNC<- N_FUNC %>% mutate_at(c("F_ATIVOS","F_FERIAS"),
                              as.numeric)

N_FUNC<- N_FUNC[,c(3,1,2)]
N_FUNC

N_FUNC$VALOR<- apply(N_FUNC[,c(2:3)],1, sum)

N_FUNC<- N_FUNC[,c(-2,-3)]

N_FUNC$DESCRICAO<- rep('QUANT FUNCIONARIOS')


N_FUNC<- aggregate(VALOR ~ TIPO+DESCRICAO, data = N_FUNC, 
                 FUN = sum, na.rm = TRUE)

N_FUNC<- N_FUNC[,c(1,3,2)]

N_FUNC<-  N_FUNC[order(N_FUNC$VALOR),]
N_FUNC

#****************************************************
# ABSENTEISMO
EVENT_FALTAS<- ('72')

FALTAS<-  folha[folha$EVENTO %in% EVENT_FALTAS & folha$TIPO=='TOTALIZACAO EMPRESA',]

FALTAS<- FALTAS[,c(8,2,4)]

FALTAS$HORAS<- as.numeric(FALTAS$HORAS)
FALTAS<- mutate(.data = FALTAS, HORAS = HORAS*-1)
                
FALTAS


EVENT_ATEST<- ('10')

ATESTATO<-  folha[folha$EVENTO %in% EVENT_ATEST & folha$TIPO=='TOTALIZACAO EMPRESA',]

ATESTADO<- ATESTATO[,c(8,2,4)]

ATESTADO$HORAS<- as.numeric(ATESTADO$HORAS)


ATESTADO

ABSENTEISMO <- rbind(FALTAS, ATESTADO)
ABSENTEISMO$DESCRICAO<- 'ABSENTEISMO'

ABSENTEISMO<- aggregate(HORAS ~ TIPO+DESCRICAO, data = ABSENTEISMO, 
                        FUN = sum, na.rm = TRUE)

colnames(ABSENTEISMO)<- c('TIPO','DESCRICAO','VALOR')


#***********************
# HORAS EXTRAS SOBRE HORAS NORMAIS

EVENT_HORAS_NORM<- ('1')

N_HORAS_NORMAIS<- folha[folha$EVENTO %in% EVENT_HORAS_NORM & folha$TIPO=='TOTALIZACAO EMPRESA',]

N_HORAS_NORMAIS<- N_HORAS_NORMAIS[,c(8,2,4)]
N_HORAS_NORMAIS$DESCRICAO<- 'QTE_HORAS_NORMAIS'

N_HORAS_NORMAIS

EVENTOS_EXTRA<- c('3','16','107') 

N_HORAS_EXTRAS<- folha[folha$EVENTO %in% EVENTOS_EXTRA & 
                        folha$TIPO=='TOTALIZACAO EMPRESA',]

N_HORAS_EXTRAS
N_HORAS_EXTRAS<- N_HORAS_EXTRAS[,c(8,2,4)]
N_HORAS_EXTRAS$DESCRICAO<- 'QTE_HORAS_EXTRAS'

N_HORAS_EXTRAS<- aggregate(HORAS ~ TIPO+DESCRICAO, data = N_HORAS_EXTRAS, 
                        FUN = sum, na.rm = TRUE)

N_HORAS_NORMAIS_N_HORAS_EXTRAS<- rbind(N_HORAS_NORMAIS,N_HORAS_EXTRAS)
HORAS_NORMAIS_SOBRE_HORAS_EXTRAS<- data.frame(N_HORAS_NORMAIS_N_HORAS_EXTRAS)
colnames(HORAS_NORMAIS_SOBRE_HORAS_EXTRAS)<- c('TIPO','DESCRICAO', 'VALOR')

#*****************ENCARGOS INSS*****************
ENCARGOS<- c('310','311')

ENCARGOS_INSS <- folha[folha$EVENTO %in% ENCARGOS,]
(ENCARGOS_INSS)

ENCARGOS_INSS<- ENCARGOS_INSS[,c(8,2,7)]

colnames(ENCARGOS_INSS)<- c('TIPO','DESCRICAO','VALOR')

ENCARGOS_INSS<- aggregate(VALOR ~ TIPO+DESCRICAO, 
                          data = ENCARGOS_INSS, 
                   FUN = sum, na.rm = TRUE)

ENCARGOS_INSS<- ENCARGOS_INSS[ENCARGOS_INSS$TIPO=='TOTALIZACAO EMPRESA',]

#***********************************************

#***********************
# CUSTO RESCISÃO
EVENT_RESC<- c('300','3','16','107','7','8','9','15','19','32','34','37',
                '39','33','47','160','164', '31','35')

DESLIGADOS<- folha[folha$EVENTO %in% EVENT_RESC,]

DESLIGADOS<- aggregate(DESLIGADOS ~ EVENTO+DESCRICAO+TIPO, data = DESLIGADOS, 
                       FUN = max, na.rm = TRUE)

DESLIGADOS


`%not_in%` <- purrr::negate(`%in%`)
DEDUCAO_DESLIGADOS<- DESLIGADOS[DESLIGADOS$EVENTO %not_in% '300',]
colnames(DEDUCAO_DESLIGADOS)<- c('EVENTO','DESCRICAO','TIPO','DEDUCAO')

DEDUCAO_DESLIGADOS
DEDUCAO_DESLIGADOS<- aggregate(DEDUCAO ~ TIPO, data = DEDUCAO_DESLIGADOS, 
                           FUN = sum, na.rm = TRUE)

DEDUCAO_DESLIGADOS

EVENT_PROVENTOS<- c('300')
CUSTO_PESSOAL_DESL<- DESLIGADOS[DESLIGADOS$EVENTO %in% EVENT_PROVENTOS,]
CUSTO_PESSOAL_DESL
CUSTO_PESSOAL_DESL<- left_join(CUSTO_PESSOAL_DESL,DEDUCAO_DESLIGADOS)
CUSTO_PESSOAL_DESL<- CUSTO_PESSOAL_DESL[,-1]

CUSTO_PESSOAL_DESL<- transform(CUSTO_PESSOAL_DESL,VALOR=DESLIGADOS - DEDUCAO)
CUSTO_PESSOAL_DESL$DEDUCAO<- NULL
CUSTO_PESSOAL_DESL$DESLIGADOS<- NULL

CUSTO_PESSOAL_DESL<- CUSTO_PESSOAL_DESL[,c(2,1,3)]
CUSTO_PESSOAL_DESL$DESCRICAO<- rep('CUSTO_COM_RESCISAO')
CUSTO_RESCISAO<- CUSTO_PESSOAL_DESL[order(CUSTO_PESSOAL_DESL$VALOR),]
CUSTO_RESCISAO

NOME_DRIVE

MES<- NOME_DRIVE %>%str_extract(pattern = "\\d+/\\d+")
#%>%str_sub(1,2)

MES

DRIVE<- NOME_DRIVE %>% 
  str_extract(pattern = "[0-9]+-[A-Za-z]+ [A-Za-z]+")

DRIVE


CUSTO_PESSOAL$DRIVE<- rep(DRIVE)
CUSTO_PESSOAL$MES<- rep(MES)



HORAS_EXTRAS$DRIVE<- rep(DRIVE)
HORAS_EXTRAS$MES<- rep(MES)


CUSTO_RESCISAO$DRIVE<- rep(DRIVE)
CUSTO_RESCISAO$MES<- rep(MES)


N_FUNC$DRIVE<- rep(DRIVE)
N_FUNC$MES<- rep(MES)  


ABSENTEISMO$DRIVE<- rep(DRIVE)
ABSENTEISMO$MES<- rep(MES)

HORAS_NORMAIS_SOBRE_HORAS_EXTRAS$DRIVE<-rep(DRIVE)
HORAS_NORMAIS_SOBRE_HORAS_EXTRAS$MES<- rep(MES)

#****
ENCARGOS_INSS$DRIVE<- rep(DRIVE)
ENCARGOS_INSS$MES<- rep(MES)

TOTAL_PROVENTOS$DRIVE<- rep(DRIVE)
TOTAL_PROVENTOS$MES<- rep(MES)

CUSTO_PESSOAL
HORAS_EXTRAS
CUSTO_RESCISAO
N_FUNC
ABSENTEISMO
HORAS_NORMAIS_SOBRE_HORAS_EXTRAS

ENCARGOS_INSS
DRO_FOLHA<- rbind(CUSTO_PESSOAL,HORAS_EXTRAS,CUSTO_RESCISAO,N_FUNC,ABSENTEISMO,
                  HORAS_NORMAIS_SOBRE_HORAS_EXTRAS, ENCARGOS_INSS, TOTAL_PROVENTOS)
#****
dim(DRO_FOLHA)

openxlsx::write.xlsx(DRO_FOLHA,paste0(letra, 
                                  ':/R_STUDIO/Folha/DRO_FOLHA_DRIVE/DRO_FOLHA_CAP1.xlsx'))


setwd(paste0(letra,':/R_STUDIO/Folha'))
rm(list = ls(all.names = TRUE))

