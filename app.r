#Aplikasi Peramalan menggunakan Model Exponential Smoothing
options(warn = -1) #agar warning tidak muncul di console

#Library yang perlu diaktifkan
library(shiny)
library(shinydashboard)
library(forecast)
library(ggplot2)
library(shinycssloaders)
library(readxl)
library(writexl)

ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(titleWidth = 300, title = "Exponential Smoothing"),
  
  dashboardSidebar(
    width = 300,
    sidebarMenu(
      menuItem("Home", tabName = "home", icon = icon("home")),
      menuItem("Deskripsi Data", tabName = "tab_desc", icon = icon("table")),
      menuItem("Dekomposisi Data", tabName = "tab_decomp", 
               icon = icon("scissors")),
      menuItem("Analisis Peramalan", tabName = "tab_forecast", 
               icon = icon("chart-line")),
      
      tags$hr(),
      fileInput("file1", "Unggah Data (Excel)", accept = ".xlsx"),
      helpText("Pastikan data dengan nama kolom (Value)!"),
      
      #Pilihan Metode
      selectInput("method", "Pilih Metode:", 
                  choices = c("Simple Exponential Smoothing" = "ses",
                              "Holt's Method (Trend)" = "holt",
                              "Holt-Winters Additive" = "hw_add",
                              "Holt-Winters Multiplicative" = "hw_mult")),
      
      #Fitur Otomatis
      checkboxInput("auto_param", "Optimalisasi Otomatis (Auto)", value = FALSE),
      
      #Parameter Manual (Hanya muncul jika Auto tidak dicentang)
      conditionalPanel(
        condition = "input.auto_param == false",
        sliderInput("alpha", "Alpha (Level):", 0.01, 1, 0.5),
        
        conditionalPanel(
          condition = "input.method == 'holt' || input.method == 'hw_add' || input.method == 'hw_mult'",
          sliderInput("beta", "Beta (Trend):", 0.01, 1, 0.05)
        ),
        
        conditionalPanel(
          condition = "input.method == 'hw_add' || input.method == 'hw_mult'",
          sliderInput("gamma", "Gamma (Seasonal):", 0.01, 1, 0.3)
        )
      ),
      
      #Input periode Time Series
      numericInput("periode", "Periode Time Series:", value = 12, min = 1),
      
      #Input periode Peramalan
      numericInput("h_period", "Periode Peramalan:", value = 12, min = 1)
    )
  ),
  
  dashboardBody(
    tabItems(
      
      #Home
      tabItem(
        tabName = "home",
        fluidRow(
          box(
            title = "Tentang Dashboard",
            status = "primary",
            solidHeader = TRUE,
            width = 12,
            
            h1(strong("Dashboard Forecasting Time Series dengan Metode Exponential Smoothing")),
            br(),
            
            tags$p(
              style = "text-align: justify;",
              "Dashboard ini digunakan untuk melakukan analisis dan peramalan data time series menggunakan metode Exponential Smoothing. Model yang tersedia meliputi Simple Exponential Smoothing, Holt's Method (Trend), Holt-Winters Additive, dan Holt-Winters Multiplicative. Setiap metode digunakan sesuai dengan karakteristik data time series yang dimiliki. Simple Exponential Smoothing digunakan untuk data tanpa tren dan musiman, Holt's Method digunakan untuk data yang memiliki tren, sedangkan Holt-Winters Additive dan Multiplicative digunakan untuk data yang memiliki pola musiman."
            ),
            br(),
            
            tags$p(
              style = "text-align: justify;",
              "Pengguna dapat mengunggah data dalam format Excel, menentukan periode time series, memilih metode peramalan, mengatur parameter alpha, beta, dan gamma secara manual maupun otomatis, serta melihat hasil peramalan dalam bentuk grafik, tabel, dan ukuran akurasi seperti RMSE, MAE, dan MAPE. Selain itu, dashboard juga menyediakan fitur dekomposisi data untuk membantu pengguna memahami pola tren, musiman, dan residual pada data sebelum melakukan peramalan. Hasil peramalan juga dapat diunduh dalam format Excel sehingga memudahkan pengguna dalam melakukan analisis lanjutan dan penyusunan laporan."
            ),
            br(),
            
            tags$hr(),
            
            h3("Informasi Penyusun"),
            p(strong("Nama"), ": Bintang Vandini"),
            p(strong("NIM"), ": 24050123130106"),
            p(strong("Departemen"), ": Departemen Statistika"),
            p(strong("Fakultas"), ": Fakultas Sains dan Matematika"),
            p(strong("Universitas"), ": Universitas Diponegoro"),
            p(strong("Tahun"), ": 2026")
          )
        )
      ),
      
      #Deskripsi Data
      tabItem(tabName = "tab_desc",
              fluidRow(
                valueBoxOutput("jumlah_data", width = 3),
                valueBoxOutput("min_box", width = 3),
                valueBoxOutput("max_box", width = 3),
                valueBoxOutput("mean_box", width = 3)
              ),
              fluidRow(
                box(title = "Summary Data", status = "info", 
                    solidHeader = TRUE, width = 12,
                    verbatimTextOutput("summary_data"))
              )
      ),
      
      #Analisis Peramalan
      tabItem(tabName = "tab_forecast",
              fluidRow(
                box(title = "Hasil Peramalan", status = "primary", 
                    solidHeader = TRUE, width = 12,
                    plotOutput("forecastPlot") %>% withSpinner(),
                    br(),
                    tableOutput("forecastTable"),
                    br(),
                    downloadButton("download_forecast", "Download Hasil Forecast"),
                    br(), br(),
                    h4("Ringkasan Hasil Model"),
                    verbatimTextOutput("model_summary"))
              ),
              fluidRow(
                valueBoxOutput("rmse_box", width = 4),
                valueBoxOutput("mae_box", width = 4),
                valueBoxOutput("mape_box", width = 4)
              ),
              fluidRow(
                valueBoxOutput("alpha_box", width = 4),
                valueBoxOutput("beta_box", width = 4),
                valueBoxOutput("gamma_box", width = 4)
              )
      ),
      
      #Dekomposisi Data
      tabItem(tabName = "tab_decomp",
              fluidRow(
                box(title = "Dekomposisi Data", status = "info", 
                    solidHeader = TRUE,width = 12,
                    plotOutput("decompPlot") %>% withSpinner())
              )
      )
    )
  )
)

server <- function(input, output) {
  
  #Data Reactive
  data_ts <- reactive({
    req(input$file1)
    df <- read_xlsx(input$file1$datapath)
    
    validate(
      need("Value" %in% names(df), "Kolom harus bernama 'Value'")
    )
    
    ts(df$Value, frequency = input$periode) 
  })
  
  #Logika Pemodelan
  model_res <- reactive({
    req(data_ts())
    
    if(input$method %in% c("hw_add", "hw_mult")) {
      validate(
        need(input$periode > 1, "Periode Time Series harus lebih dari 1 untuk metode Holt-Winters")
      )
    }
    
    if(input$method == "hw_mult") {
      validate(
        need(all(data_ts() > 0), "Data harus bernilai positif untuk Holt-Winters Multiplicative")
      )
    }
    
    #Setting Parameter (NULL jika ingin Auto)
    a <- if(input$auto_param) NULL else input$alpha
    b <- if(input$auto_param) NULL else input$beta
    g <- if(input$auto_param) NULL else input$gamma
    
    #Pemilihan Fungsi berdasarkan Input
    if(input$method == "ses") {
      ses(data_ts(), alpha = a, h = input$h_period)
      
    } else if(input$method == "holt") {
      holt(data_ts(), alpha = a, beta = b, h = input$h_period)
      
    } else if(input$method == "hw_add") {
      hw(data_ts(), seasonal = "additive", alpha = a, beta = b, gamma = g, h = input$h_period)
      
    } else {
      hw(data_ts(), seasonal = "multiplicative", alpha = a, beta = b, gamma = g, h = input$h_period)
    }
  })
  
  #Plot Forecast
  output$forecastPlot <- renderPlot({
    autoplot(model_res()) +
      autolayer(fitted(model_res()), series = "Fitted") +
      theme_minimal() + 
      labs(y="Data", x="Time") +
      theme(legend.position = "bottom")
  })
  
  #Plot Dekomposisi
  output$decompPlot <- renderPlot({
    req(data_ts())
    
    validate(
      need(length(data_ts()) > input$periode * 2,
           "Data terlalu pendek untuk dekomposisi")
    )
    
    autoplot(decompose(data_ts())) + theme_minimal()
  })
  
  #Summary Data
  output$summary_data <- renderPrint({
    summary(data_ts())
  })
  
  #Valuebox Deskripsi Data
  
  #Jumlah Data
  output$jumlah_data <- renderValueBox({
    req(data_ts())
    valueBox(length(data_ts()), "Jumlah Data", icon = icon("database"), color = "blue")
  })
  
  #Min
  output$min_box <- renderValueBox({
    req(data_ts())
    valueBox(round(min(data_ts()),2), "Minimum", icon = icon("arrow-down"), color = "green")
  })
  
  #Max
  output$max_box <- renderValueBox({
    req(data_ts())
    valueBox(round(max(data_ts()),2), "Maksimum", icon = icon("arrow-up"), color = "red")
  })
  
  #Rata-Rata (Mean)
  output$mean_box <- renderValueBox({
    req(data_ts())
    valueBox(round(mean(data_ts()),2), "Rata-rata", icon = icon("chart-line"), color = "yellow")
  })
  
  #Tabel Forecast
  output$forecastTable <- renderTable({
    fc <- model_res()
    
    data.frame(
      Periode = 1:length(fc$mean),
      Forecast = round(as.numeric(fc$mean), 2),
      Lower_80 = round(as.numeric(fc$lower[,1]), 2),
      Upper_80 = round(as.numeric(fc$upper[,1]), 2),
      Lower_95 = round(as.numeric(fc$lower[,2]), 2),
      Upper_95 = round(as.numeric(fc$upper[,2]), 2)
    )
  })
  
  #Download Forecast
  output$download_forecast <- downloadHandler(
    filename = function() {
      "hasil_forecast.xlsx"
    },
    content = function(file) {
      fc <- model_res()
      
      hasil <- data.frame(
        Periode = 1:length(fc$mean),
        Forecast = as.numeric(fc$mean),
        Lower_80 = as.numeric(fc$lower[,1]),
        Upper_80 = as.numeric(fc$upper[,1]),
        Lower_95 = as.numeric(fc$lower[,2]),
        Upper_95 = as.numeric(fc$upper[,2])
      )
      
      write_xlsx(hasil, path = file)
    }
  )
  
  #Ringkasan Hasil Model
  output$model_summary <- renderPrint({
    summary(model_res())
  })
  
  #Valuebox Akurasi
  
  #RMSE
  output$rmse_box <- renderValueBox({
    acc <- accuracy(model_res())
    valueBox(round(acc[1,"RMSE"],2), "RMSE", 
             icon = icon("triangle-exclamation"), color = "red")
  })
  
  #MAE
  output$mae_box <- renderValueBox({
    acc <- accuracy(model_res())
    valueBox(round(acc[1,"MAE"],2), "MAE", 
             icon = icon("chart-line"), color = "yellow")
  })
  
  #MAPE
  output$mape_box <- renderValueBox({
    acc <- accuracy(model_res())
    valueBox(paste0(round(acc[1,"MAPE"],2), "%"), "MAPE", 
             icon = icon("percent"), color = "green")
  })
  
  #Valuebox Parameter Model
  
  #Alpha
  output$alpha_box <- renderValueBox({
    req(model_res())
    
    val <- tryCatch({
      if(!is.null(model_res()$model$par["alpha"])) {
        round(as.numeric(model_res()$model$par["alpha"]), 3)
      } else if(!is.null(model_res()$model$alpha)) {
        round(as.numeric(model_res()$model$alpha), 3)
      } else {
        round(input$alpha, 3)
      }
    }, error = function(e) {
      round(input$alpha, 3)
    })
    
    valueBox(val, "Alpha (Level)", icon = tags$i(HTML("&alpha;"), style = "font-size:40px; font-style:normal;"), color = "light-blue")
  })
  
  #Beta
  output$beta_box <- renderValueBox({
    req(model_res())
    
    val <- tryCatch({
      if(input$method %in% c("holt","hw_add","hw_mult")) {
        if(!is.null(model_res()$model$par["beta"])) {
          round(as.numeric(model_res()$model$par["beta"]), 3)
        } else if(!is.null(model_res()$model$beta)) {
          round(as.numeric(model_res()$model$beta), 3)
        } else {
          round(input$beta, 3)
        }
      } else {
        "-"
      }
    }, error = function(e) {
      if(input$method %in% c("holt","hw_add","hw_mult")) round(input$beta, 3) else "-"
    })
    
    valueBox(val, "Beta (Trend)", icon = tags$i(HTML("&beta;"), style = "font-size:40px; font-style:normal;"), color = "olive")
  })
  
  #Gamma
  output$gamma_box <- renderValueBox({
    req(model_res())
    
    val <- tryCatch({
      if(input$method %in% c("hw_add","hw_mult")) {
        if(!is.null(model_res()$model$par["gamma"])) {
          round(as.numeric(model_res()$model$par["gamma"]), 3)
        } else if(!is.null(model_res()$model$gamma)) {
          round(as.numeric(model_res()$model$gamma), 3)
        } else {
          round(input$gamma, 3)
        }
      } else {
        "-"
      }
    }, error = function(e) {
      if(input$method %in% c("hw_add","hw_mult")) round(input$gamma, 3) else "-"
    })
    
    valueBox(val, "Gamma (Seasonal)", icon = tags$i(HTML("&gamma;"), style = "font-size:40px; font-style:normal;"), color = "purple")
  })
}

shinyApp(ui, server)
