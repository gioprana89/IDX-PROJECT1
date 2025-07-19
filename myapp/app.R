

#library(webr)


#library(plyr)
library(readxl)
#library(openxlsx)
#library(data.table)
library(shiny)
#library(shinyAce)
#source("chooser.R")
library(scales)
library(ggplot2)

library(dplyr)
#library(lavaan)

#library(mnormt)
#library(curl)
#library(plspm)


########################################
########UI (User Interface)#############
########################################

modul_data_IDX_ui <- function(id) {
  
  
  
  ns <- NS(id)
  
  fluidPage(
    
    
    
    
    
    
    
    tabsetPanel(
      
      tabPanel(title = tags$h5( tags$img(src = "dataset.png", width = "30px"), 'Data Selection'),
               
               
               
               
               
               uiOutput(ns("buka_pemilihan_informasi")),
               
               
               br(),
               
               uiOutput(ns("pilih_sektor")), 
               
               
               
               br(),
               
               uiOutput(ns("buka_nama_subsektor")), 
               
               
               br(),
               
               uiOutput(ns("pilih_tahun")), 
               
               
               
               
               
               
               br(),
               
               uiOutput(ns("pilih_perusahaan")), 
               
               
               
               
               
               
               br(),
               
               
               
               br()
               
      ), #tabpanel Data Selection
    
    
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      tabPanel(title = tags$h5( tags$img(src = "statistics.png", width = "30px"), 'Statistics'),
               
              
               
                    
                  h1("Financial Data & Ratio",
                     style="color:red;
                     text-align:center;
                     font-size:35px;"         ),
               
               
               
               br(),
               
               
               DT::DTOutput(ns("buka_data")),
               
               
               
               br(),
               br(),
               br(),
               
               
               
               
               
               
               
               
               
               
               
               
               
               h1("Number of Data",
                  style="color:red;
                     text-align:center;
                     font-size:35px;"         ),
               
               
               
               br(),
               
               
               DT::DTOutput(ns("buka_data_informasi_jumlah_sektor")),
               
               
               
               
               
               
               
               
               br(),
               
               br(),
               
               br(),
               
   
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               h1("Descriptive Statistics",
                  style="color:red;
                     text-align:center;
                     font-size:35px;"         ),
               
               
               
               br(),
               
               
               uiOutput(ns("buka_variabel_deskriptif_bivariat")), 
        
               br(),
               
               
               DT::DTOutput(ns("buka_data_deskriptif_bivariat")),
               
               
               
               
               
               
               
               br(),
               
               
               br(),
               
               
               br(),
               
               
          
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               
               h1("Correlation Matrix",
                  style="color:red;
                     text-align:center;
                     font-size:35px;"         ),
               
               
               
               br(),
               
               
               
               uiOutput(ns("buka_pemilihan_variabel_numerik_untuk_matriks_korelasi")),
               
               
               br(),
               
               
               
               plotOutput(ns("grafik_batang_matriks_korelasi"), width = "1400px", height = "800px" ),
               
               
               
               
               
               
               
               
               
               
               
               
               
               
                br()
               
               
               
               
               
               
               
      ), #End of tabpanel Statistics
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      tabPanel(title = tags$h5( tags$img(src = "excel.png", width = "30px"), 'Print Data to Excel'),
               
               
               
               
               h1("Print Data to Excel",
                  style="color:red;
                     text-align:center;
                     font-size:35px;"         ),
               
               
               
               
       verbatimTextOutput(ns("cetak_data")),
               
               
               
               
               
               
               
               
               
               
               
               br()
               

               
      ) #End of tabpanel print data to excel
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
    ), #End of tabsetPanel
    
    
    
    
    
    
    
    
    
    
    br()
    
  ) #Akhir Fluidpage
  
  
} #Akhir dari modul_data_IDX_ui

#Akhir dari modul_data_IDX_ui
#Akhir dari modul_data_IDX_ui
#Akhir dari modul_data_IDX_ui
#Akhir dari modul_data_IDX_ui











































































########################################
################Server##################
########################################



modul_data_IDX_server <- function(input, output, session) {
  
  
  
  
  
  ###########Fungsi untuk menyeleksi data berdasarkan Variabel, Sektor, Subsektor, Tahun dan Perusahaan
  ###########Fungsi untuk menyeleksi data berdasarkan Variabel, Sektor, Subsektor, Tahun dan Perusahaan
  ###########Fungsi untuk menyeleksi data berdasarkan Variabel, Sektor, Subsektor, Tahun dan Perusahaan
  
  ############################################
  ############Pemilihan Variabel##############
  ############################################
  
  
  nama_variabel <- function()
  {
    
    dat <- read_xlsx("DATA IDX.xlsx")
    dat <- as.data.frame(dat)
    
    nama <- colnames(dat)
    
    return(nama)
    
    
  }
  
  output$buka_pemilihan_informasi <- renderUI({
    
    
    
    
    
    
    checkboxGroupInput(session$ns("terpilih_variabel"), 
                       label="Select Variables:", choices = c(nama_variabel()), 
                       selected=c("Sector", "Sub Industry", "Code", "Stock Name",   
                                  "Profit for the Period", "Profit attr.to owner's", "EPS", "Book Value", "PER", "PBV", 
                                  "DER", "ROA", "ROE", "NPM", "Data Time"), inline = TRUE)
    
    
    
  })
  
  
  
  
  
  
  
  
  
  ###############Pemilihan Sektor##################
  ###############Pemilihan Sektor##################  
  ###############Pemilihan Sektor##################  
  
  nama_sektor <- function()
  {
    
    dat <- read_xlsx("DATA IDX.xlsx")
    dat <- as.data.frame(dat)
    
    
    nama_sektor <- dat[, c("Sector")]
    
    nama_sektor <- as.character(nama_sektor)
    nama_sektor <- as.factor(nama_sektor)
    
    nama_sektor <- levels(nama_sektor)
    
    return(nama_sektor)
    
  }
  
  
  output$pilih_sektor <- renderUI({
    
    
    
    
    
    
    checkboxGroupInput(session$ns("terpilih_sektor"), 
                       label="Select Sectors:", choices = c(nama_sektor()), 
                       selected=c( nama_sektor() ), inline = TRUE)
    
    
    
  })
  
  
  
  
  
  
  
  
  #################################################
  ###############Pemilihan Subsektor###############
  ###############Pemilihan Subsektor###############  
  
  nama_fungsi_terpilih_subsektor <- function()
  {
    
    
    
    dat <- read_xlsx("DATA IDX.xlsx")
    dat <- as.data.frame(dat)
    
    
    
    
    terpilih_sektor <- input$terpilih_sektor
    terpilih_sektor <- as.character(terpilih_sektor)
    
    
    
    full_sektor <- dat[,"Sector"]
    
    
    
    
    
    indeks <- full_sektor %in% terpilih_sektor 
    indeks <- which(indeks==TRUE)
    
    dat2 <- dat[c(indeks),]
    
    
    nama_subsektor <- dat2[, "Sub Industry"]
    
    nama_subsektor <- as.character(nama_subsektor)
    nama_subsektor <- as.factor(nama_subsektor)
    
    nama_subsektor <- levels(nama_subsektor)
    
    return(nama_subsektor)
    
    
    
    
  }
  
  
  
  #################
  
  
  output$buka_nama_subsektor <- renderUI({
    
    
    
    
    
    
    checkboxGroupInput(session$ns("terpilih_subsektor"), 
                       label="Select Subsectors:", choices = c(nama_fungsi_terpilih_subsektor()), 
                       selected=c( nama_fungsi_terpilih_subsektor()    ), inline = TRUE)
    
    
    
  })
  
  
  
  
  
  
  
  
  
  #############################################
  ################Pemilihan Tahun##############
  #############################################
  
  
  output$pilih_tahun <- renderUI({
    
    
    
    
    
    
    checkboxGroupInput(session$ns("terpilih_tahun"), 
                       label="Select Years:", choices = c(2021, 2022, 2023, 2024), 
                       selected=c(2021, 2022, 2023, 2024), inline = TRUE)
    
    
    
  })
  
  
  
  
  
  
  
  
  ######################################################
  ####################Pemilihan Perusahaan##############
  ######################################################
  
fungsi_nama_perusahaan <- function()
{
  
  
  dat <- read_xlsx("DATA IDX.xlsx")
  dat <- as.data.frame(dat)
  
  
  
  #####Seleksi 1: Pemilihan Variabel#######
  
  nama <- colnames(dat)
  
  terpilih_variabel <- input$terpilih_variabel
  
  dat_baru <- dat[c(terpilih_variabel)]
  
  
  
  
  
  
  #####Seleksi 2: Pemilihan Sektor#######  
  
  ###############
  ###############
  
  dat <- dat_baru
  
  terpilih_sektor <- input$terpilih_sektor
  terpilih_sektor <- as.character(terpilih_sektor)
  
  full_sektor <- dat[,"Sector"]
  
  indeks <- full_sektor %in% terpilih_sektor 
  indeks <- which(indeks==TRUE)
  
  dat2 <- dat[c(indeks),]
  
  
  
  
  #####Seleksi 3: Pemilihan Subsektor#######  
  
  terpilih_subsektor <- input$terpilih_subsektor
  terpilih_subsektor <- as.character(terpilih_subsektor)
  
  full_subsektor <- dat2[,"Sub Industry"]
  
  indeks <- full_subsektor %in% terpilih_subsektor 
  indeks <- which(indeks==TRUE)
  
  dat3 <- dat2[c(indeks),]
  
  
  
  
  
  #####Seleksi 4: Tahun#######  
  
  full_tahun = dat3[,c("Data Time")]
  
  
  terpilih_tahun <- input$terpilih_tahun
  terpilih_tahun <- as.numeric(terpilih_tahun)
  
  
  indeks <- full_tahun %in% terpilih_tahun 
  indeks <- which(indeks==TRUE)
  
  dat4 <- dat3[c(indeks),]
  
  
  
  
  
  dat <- dat4
  nama_perusahaan <- dat[, "Code"]
  
  nama_perusahaan <- as.factor(nama_perusahaan)
  
  nama_perusahaan <- levels(nama_perusahaan)
  
  
  return(nama_perusahaan)
  
}
  
  
  
  ################
  
  
  
  
  output$pilih_perusahaan <- renderUI({
    
    
    
    
    
    
    checkboxGroupInput(session$ns("terpilih_perusahaan"), 
                       label="Select Companies:", choices = c(fungsi_nama_perusahaan()), 
                       selected=c( fungsi_nama_perusahaan()    ), inline = TRUE)
    
    
    
  })
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  #####################
  #####################
  
  
  
  fungsi_seleksi_data <- function()
  {
    
    dat <- read_xlsx("DATA IDX.xlsx")
    dat <- as.data.frame(dat)
    
    
    
    #####Seleksi 1: Pemilihan Variabel#######
    
    nama <- colnames(dat)
    
    terpilih_variabel <- input$terpilih_variabel
    
    dat_baru <- dat[c(terpilih_variabel)]
    
    
    
    
    
    
    #####Seleksi 2: Pemilihan Sektor#######  
    
    ###############
    ###############
    
    dat <- dat_baru
    
    terpilih_sektor <- input$terpilih_sektor
    terpilih_sektor <- as.character(terpilih_sektor)
    
    full_sektor <- dat[,"Sector"]
    
    indeks <- full_sektor %in% terpilih_sektor 
    indeks <- which(indeks==TRUE)
    
    dat2 <- dat[c(indeks),]
    
    
    
    
    #####Seleksi 3: Pemilihan Subsektor#######  
    
    terpilih_subsektor <- input$terpilih_subsektor
    terpilih_subsektor <- as.character(terpilih_subsektor)
    
    full_subsektor <- dat2[,"Sub Industry"]
    
    indeks <- full_subsektor %in% terpilih_subsektor 
    indeks <- which(indeks==TRUE)
    
    dat3 <- dat2[c(indeks),]
    
    
    
    
    
    #####Seleksi 4: Tahun#######  
    
    full_tahun = dat3[,c("Data Time")]
    
    
    terpilih_tahun <- input$terpilih_tahun
    terpilih_tahun <- as.numeric(terpilih_tahun)
    
    
    indeks <- full_tahun %in% terpilih_tahun 
    indeks <- which(indeks==TRUE)
    
    dat4 <- dat3[c(indeks),]
    
    
    ###############Seleksi 5: Perusahaan atau Code###############
    
    
    
    #print(dat4)
    #print(dat4)
    
    full_perusahaan = dat4[,c("Code")]
    
    
    terpilih_perusahaan <- input$terpilih_perusahaan
    terpilih_perusahaan <- as.character(terpilih_perusahaan)
    
    
    indeks <- full_perusahaan %in% terpilih_perusahaan
    indeks <- which(indeks==TRUE)

    
    dat5 <- dat4[c(indeks),]
    
    
    
    return(dat5)
    
    
    
    
  }
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  #####################
  #####################
  #####################
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  ##################
  
  
  
  
  
  output$buka_data <- DT::renderDT({
    
    
    p <- fungsi_seleksi_data()
    
    print(p)
    
    
    
    
  })
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  #####################
  #####################
  
  
  
  
  
  
  
  
  
  
  output$buka_data_informasi_jumlah_sektor <- DT::renderDT({
    
    p <- fungsi_seleksi_data()
    
        grup <- group_by(p, Sector)
    
    dframe_frekuensi_sektor <- grup %>% summarise(
      
      Jumlah = n()
      
      
    )
    
    
    dframe_frekuensi_sektor <- as.data.frame(dframe_frekuensi_sektor)
    
    
   print(dframe_frekuensi_sektor)
    
    
    
  })
    
    
    
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   #################Deskriptif Bivariat
   
   fungsi_nama_variabel_numerik_deskriptif_univariat <- function()
   {
     
     nama <- c("Profit for the Period", "Profit attr.to owner's", "EPS", "Book Value", "PER", "PBV", "DER", "ROA", "ROE", "NPM")
     
     return(nama)
     
     
   }
   
   
   
   output$buka_variabel_deskriptif_bivariat <- renderUI({
     
     
     
     
     radioButtons(session$ns("terpilih_var_numerik_bivariat"), 
                  "Choose Numeric Variable:",
                  choices = c( fungsi_nama_variabel_numerik_deskriptif_univariat()   ),
                  selected = "ROA", inline = TRUE)
   
   

     
   })
   
   
   
   #############
   
   
   
   
   
   output$buka_data_deskriptif_bivariat <- DT::renderDT({
     
     
     p <- fungsi_seleksi_data()
     
     
     
     terpilih_var_numerik_bivariat <- input$terpilih_var_numerik_bivariat
     terpilih_var_numerik_bivariat <- as.character(terpilih_var_numerik_bivariat)
     
     
 
     
     nama_numerik <- c(terpilih_var_numerik_bivariat)
     
     dat_deskriptif <- p[c("Sector", nama_numerik )]
     
     
     
     
     library(dplyr)
     
     colnames(dat_deskriptif)[2] = c("value")
     
     
     
     grup <- group_by(dat_deskriptif, Sector)
     
     dframe_deskriptif <- grup %>% summarise(
       
       RataRata = mean(value),
       StandarDeviasi = sd(value),
       Minimum = min(value),
       Maksimum = max(value),
       jumlah = n()
       
       
     )
     
     
     dframe_deskriptif <- as.data.frame(dframe_deskriptif)
     
     
     print(dframe_deskriptif)
     
     
     
     
   })
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
    
    
  
  
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   #################
   
   
   
   #################Matriks Korelasi
   
   fungsi_nama_variabel_numerik_untuk_matriks_korelasi <- function()
   {
     
     nama <- c("Profit for the Period", "Profit attr.to owner's", "EPS", "Book Value", "PER", "PBV", "DER", "ROA", "ROE", "NPM")
     
     return(nama)
     
     
   }
   
   
   
   output$buka_pemilihan_variabel_numerik_untuk_matriks_korelasi <- renderUI({
     
     
     
     
     checkboxGroupInput(session$ns("terpilih_variabel_numerik_untuk_matriks_korelasi"), 
                  "Choose Numeric Variable:",
                  choices = c( fungsi_nama_variabel_numerik_untuk_matriks_korelasi()   ),
                  selected =  fungsi_nama_variabel_numerik_untuk_matriks_korelasi() , inline = TRUE)
     
     
     
     
   })
   
   
   
   
   
   #######################
   
   
   output$grafik_batang_matriks_korelasi <- renderPlot({
     
     p <- fungsi_kirim_grafik_matriks_korelasi()
     
     print(p)
     
     
   })
   
   
   
   
   ################
   
   
   fungsi_kirim_grafik_matriks_korelasi <- function()
   {
     
     
     p <- fungsi_seleksi_data()
     
     
     
     
     ###################
     
     
     terpilih_variabel_numerik_untuk_matriks_korelasi <- input$terpilih_variabel_numerik_untuk_matriks_korelasi
     terpilih_variabel_numerik_untuk_matriks_korelasi <- as.character(terpilih_variabel_numerik_untuk_matriks_korelasi)
     
     
     dat_matriks_korelasi <- p[c(terpilih_variabel_numerik_untuk_matriks_korelasi)]
     
     
     library(ggplot2)
     
     library(ggcorrplot)
     
     dat_matriks_korelasi <- p[c(terpilih_variabel_numerik_untuk_matriks_korelasi)]
     
     dat_matriks_korelasi
     
     corr <- round(cor(dat_matriks_korelasi), digits = 3)
     
     
     p <- ggcorrplot(corr, hc.order = TRUE, type = "lower",
                lab = TRUE)
     
     
     return(p)
     
     
   }
   
   
   
   
   
   
   
   
   
   
   
   ###################Cetak Data
   
   
   
   output$cetak_data <- renderPrint({
     
     
     p <- fungsi_seleksi_data()
     
     
     
     data_cetak <- p
     
     data_cetak <- as.data.frame(data_cetak)
     
     
     wb <- openxlsx::createWorkbook()
     
     
     hs1 <- openxlsx::createStyle(fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "italic",
                                  border = "Bottom")
     
     
     hs2 <- openxlsx::createStyle(fontColour = "#ffffff", fgFill = "#4F80BD",
                                  halign = "center", valign = "center", textDecoration = "bold",
                                  border = "TopBottomLeftRight")
     
     openxlsx::addWorksheet(wb, "Data IDX", gridLines = TRUE)
     
     openxlsx::writeDataTable(wb, "Data IDX", data_cetak, rowNames = FALSE, startRow = 2, startCol = 2, tableStyle = "TableStyleMedium21")
     
     
     
     openxlsx::openXL(wb)
     
     
   })
   
   
   
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
} #akhir dari modul_data_IDX_server

#akhir dari modul_data_IDX_server
#akhir dari modul_data_IDX_server
#akhir dari modul_data_IDX_server

















































































ui <- fluidPage(
  
  
  includeHTML("intro_home.html"),
  
  
  uiOutput("modul_data_IDX"),
  
  
  br()
  
) #Akhir dari UI











server <- function(input, output) {
  
  
  
  
  
  output$modul_data_IDX <- renderUI({
    
    
    
    #source("module//modul_data_IDX.R")
    callModule(module = modul_data_IDX_server, id = "modul_data_IDX")
    modul_data_IDX_ui(id = "modul_data_IDX")
    
    
    
  })
  
  
  
  
  
  
  
  
  
  
  
} #Akhir dari server










shinyApp(ui, server)














