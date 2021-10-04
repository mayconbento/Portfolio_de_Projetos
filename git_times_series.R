#*****************PREVISÕES TIMES SERIES  **********************************
#*****************BIBLIOTECA PROPHET****************************************

# LIMPAR VARIÁVEIS
rm(list = ls(all.names = TRUE))
getwd()

# Carregando as bibliotecas necessárias
library("gt")
library('readxl')
library('openxlsx')
library('readxl')
library('ggplot2')
library('magrittr')
library('stringr')
library('chron')
library('rJava')
library("lubridate")
library('dplyr')
library('tidyr')
library('purrr')
library('forecast')
library('xts')
library('zoo')
library('prophet')
library('matrixStats')
library('Rcpp')

# VERIFICAR DIRETÓRIO DE TRABALHO
getwd()
letra<- getwd()
letra<- str_sub(letra, start = 1, end = 1) 
dir<- paste0(letra,':/Pasta_arquivos')

# DEINIR DIREÓRIO DE TRABALHO
setwd(dir)
path <- 'arquivos-anual-2019.xlsx'

# LER MULTIPLAS ABAS DO EXCEL E FAZER  ROW BINDING
df<- path %>%
  excel_sheets() %>%
  set_names() %>%
  map_df(~ read_excel(path = path, sheet = .x, col_names = FALSE), 
         .id = "MES")
head(df)
colnames(df)<- df[2,]
str(df)

# LIMPEZA DOS DADOS E SUBSETTING
df<- df[-1:-2,]
df<- df[,c(2,3,16,19,20)]
str(df)

# ELIMINAR VALORES NA
df<- df[complete.cases(df),]

# CONTAR VALORES NA
sapply(df, function(x)sum(is.na(x)))

df_original<- df

df<- df[df$`NOME DA EMPRESA`=='EMPRESA-01',]

head(df)

# EXTRAIR NOME DA EMPRESA
DRIVE<- unique(df$`NOME DA EMPRESA`)
DRIVE

# PREPARANDO DADOS PARA O PROPHET
colnames(df)<- c('ds', 'EMPRESA','TOTAL','VALOR',
                 'EC')

df<- mutate_at(.tbl = df, .vars = c('TOTAL',
                                    'VALOR', 
                                    'EC'),  .funs = as.numeric)

df$ds<- as.Date(df$ds, origin = '1899-12-30', format = '%d/%m/%Y')

str(df)

# AGRUPAMENTO GROUP-BY
df<- df %>%
  group_by(ds) %>%
  summarise(`TOTAL` =sum(`TOTAL`),
            VALOR = sum(VALOR),
            `EC` = sum(`EC`))


head(df)
tail(df)

# SELECIONAR ULTIMAS 120 LINHAS
df<- xts::last(x = df, 120)

dim(df)

# EXCLUIR ULTIMO DIA
ultimo_dia<- xts::last(df$ds,1)

`%not_in%` <- purrr::negate(`%in%`)

df<- df[df$ds %not_in% ultimo_dia,]
head(df)
tail(df)

# DEFINIR DIAS A FRENTE (HORIZON)
MES<- month(last(df$ds))
DAY<- day(last(df$ds))
dias_a_frente<- days_in_month(MES)-DAY
dias_a_frente<- as.numeric(dias_a_frente)
dias_a_frente

#ggplot(df, mapping = aes(x= ds, y = VALOR))+
#  theme_minimal()+
#  geom_point()+
#  geom_line()+
#  ggtitle('VALOR versus Dia')+
#  theme(
#    plot.title = element_text(size=17, face="bold", hjust = 0.5))

df_TOTAL<- df %>%select(ds,`TOTAL)%>%
  set_colnames(c('ds','y'))

df_VALOR<- df %>%select(ds,VALOR)%>%
  set_colnames(c('ds','y'))

df_EC<- df %>%select(ds,`EC`)%>%
  set_colnames(c('ds','y'))

# INSERINDO OS FERIADOS NO MODELO
holidays <- tibble(
  holiday = c('feriado'),
  ds = as.Date(c('2020-09-07','2020-10-12', '2020-11-02', '2020-11-20',
                 '2020-12-24','2020-12-25', '2020-12-31', '2020-12-30',
                 '2021-01-01','2021-04-02','2021-04-21','2021-06-03',
                 '2021-07-09','2021-09-07')),
  lower_window = 0,
  upper_window = 0
)

#write.xlsx(df_VALOR, 'VALOR.xlsx')

# CONSTRUINDO O MODELO COM PROPHET
VALOR_model<- prophet(df_VALOR, 
                        seasonality.mode = 'multiplicative',
                        daily.seasonality = TRUE, 
                        holidays = holidays)
#,seasonality.prior.scale = 10,
#changepoint.prior.scale = 0.5)

total_model<- prophet(df_VALOR, 
                      seasonality.mode = 'multiplicative',
                      daily.seasonality = TRUE, 
                      holidays = holidays)


ec_model<- prophet(df_EC, 
                     seasonality.mode = 'multiplicative',
                     daily.seasonality = TRUE, 
                     holidays = holidays)

future<- make_future_dataframe(VALOR_model, 
                               periods = dias_a_frente, 
                               freq = 'day')
head(future)
tail(future)

VALOR_forecast<- predict(VALOR_model, future)
eco_forecast<- predict(econ_model, future)
total_forecast<- predict(total_model, future)

# Plot Forecasting VALOR
dyplot.prophet(VALOR_model, VALOR_forecast)

# Plot Forecasting TOTAL
dyplot.prophet(total_model, total_forecast)

# Plot Componentes
prophet_plot_components(VALOR_model, VALOR_forecast)

# Performance evaluation
df.cv <- cross_validation(VALOR_model, 
                          initial = 73, 
                          horizon = 15, 
                          units = 'days')

(df.cv)

# Ploting the performance (MEAN AVERAGE PERCENTAGE ERROR)
df.p <- performance_metrics(df.cv, metrics = 'mape')
(df.p)

#RESULTS CROSS VALIDATION
#horizon          mape
#1 24 hours 0.02777898583
#2 36 hours 0.02807610576
#3 48 hours 0.05799829397
#4 60 hours 0.04538060742
#5 72 hours 0.05446994455
#6 84 hours 0.03634461563

plot_cross_validation_metric(df.cv, metric = 'mape')

forecast<- bind_cols(ec_forecast$ds,total_forecast$yhat,
                     VALOR_forecast$yhat,
                     ec_forecast$yhat)

colnames(forecast)<- c('ds','TOTAL','VALOR',
                       'EC')

forecast

MES_ATUAL<- df[month(df$ds)==MES,]

FORECAST<- tail(forecast[,],dias_a_frente)

str(FORECAST)

# CONSTRUIR DATASET COM DADOS ATUAIS + FORECAST
completo<-rbind(MES_ATUAL,FORECAST)
completo$UNIDADE<- UNIDADE
completo<- completo[,c(5,1,2,3,4)]

hoje<- Sys.Date()
hoje

openxlsx::write.xlsx(completo, 
                     paste0(UNIDADE,' forecast_VALOR-',hoje,'.xlsx'))

