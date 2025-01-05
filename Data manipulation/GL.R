#Always run when opened---------------------------------------------------------
{
  options(warn = -1)
  rm(list = ls())
  cat("\014")
  gc()
  
  options(verbose = FALSE)
  options(scipen = 999)
  requiredPackages = c(
    "stringr",
    
    "foreign",
    "data.table",
    "RMySQL",
    "RCurl",
    "TTR",
    "httr",
    "gdata",
    "tableHTML",
    "textclean",
    "rvest",
    "XML",
    "Rcrawler",
    "knitr",
    "BatchGetSymbols",
    "Quandl",
    "anytime",
    "quantmod",
    "dplyr",
    "tibble",
    "base64enc",
    "lubridate",
    "curl",
    "tidyr",
    "WDI",
    "zoo",
    "ggplot2",
    "e1071",
    "tseries",
    "TeachingDemos",
    "psych",
    "lmtest",
    "rugarch",
    "Hmisc",
    "RColorBrewer",
    "broom",
    "car",
    "sandwich",
    "forecast",
    "zoo",
    "moments",
    "qrmdata",
    "qrmtools",
    "carData",
    "survival",
    "dynlm",
    "urca",
    "fUnitRoots",
    "ggpubr",
    "MASS",
    "robust",
    "readxl",
    "WRS2",
    "robustrank",
    "glue",
    "slider",
    "gamlss",
    "usethis",
    "rstudioapi",
    "RPostgres",
    "DBI",
    "haven",
    'outliers',
    "pdfetch",
    "olsrr",
    "openxlsx",
    "broom",
    "lubridate",
    "stats4",
    "matrixStats",
    "optimx",
    "vars",
    "mFilter",
    "readxl",
    "writexl",
    "GGally",
    "rmgarch",
    "lognorm",
    "tidyquant",
    "xdcclarge",
    'MTS',
    "ggcorrplot",
    "xml2",
    "wbstats",
    "beepr",
    "printr",
    "shiny",
    "shinythemes",
    "bslib",
    "rmarkdown",
    "sodium",
    "shinyjs",
    'evir'
    
  )
  
  for (p in requiredPackages) {
    print(p, quote = F)
    cat("-----------------------------------------------------------------------","\n")
    if (!require(p, character.only = TRUE))
      install.packages(p, quiet = T)
    library(p, character.only = TRUE, quietly = T)
  }
  # install.packages('pdftools', type='source')
  setwd("C:/Users/KhaiPham/OneDrive - Grant Thornton (Vietnam) Limited/Quang Khai/Mar Project/Rstudio")
  wd = getwd()
  wd
  source(paste0(wd, "/IFRC_TRAINEE_2020_LIBRARY.R"), echo = F)
  # source(paste0(wd, "/thuonglibrary.R"), echo = F)
  # source(paste0(wd, "/NCKH_LIBRARY.R"), echo = F)
  # source("/Users/phamlequangkhai/Documents/WinBlack_Mac/R/KLTN_LIBRARY.R", echo = F)
  # source("/Library/Frameworks/R.framework/Versions/4.1/Resources/library/ccgarch/dcc_estimate.R", echo = F)
  RESULT           = paste0(wd, '/RESULT')
  DATA_source      = paste0(wd, '/DATA_RAW/')
  textbox_to_table = function(dir, excel_File_Name){
    f     = extract_Textbox_From_Excel(dir, excel_File_Name)
    f1    = str_split(f, "Tên tài khoản")
    f1
    title = data.table()
    for (i in 2:length(f1[[1]]))
    {
      # i=2
      cat("\014")
      print(paste("Tai khoan", f1[[1]][i]))
      # f1[[1]][i]
      stk   = str.extract(f1[[1]][i], ": ", " -")
      stk
      ttk   = str.extract(f1[[1]][i], " - ", "- S")
      ttk
      stk   = cbind(stk, ttk)
      title = rbind(title, stk)
    }
    
  }
  work_complete    = function() {
    cat("Work complete. Press Esc to sound the fanfare!!!\n")
    on.exit(beepr::beep(4, expr = "Done everything, have fun now"))
    
    while (TRUE) {
      beep(10)
      Sys.sleep(1)
    }
  }
  #for Mac: source("/Volumes/Quang Khai/WinBlack/R/NCKH_LIBRARY.R", echo = F)
  cat("\014")
  format(Sys.Date(), "%d/%m/%Y")
  print("Hello bro")
  beep(5)
  
}

#Lay phan textbox--------------------------------------------------------------------
extract_Textbox_From_Excel <- function(dir, excel_File_Name)
{
   library(stringr)
   library(xml2)
   setwd(wd)
   file_Name_No_Ext <- str_replace(string = excel_File_Name, pattern = "\\.xls", replacement = "")
   file_Name_Zip <- paste0(file_Name_No_Ext, ".zip")
   dir.create(tmp <- paste0(wd,"/tmp"))
   file.copy(from = paste0(wd,"/DATA_RAW/",excel_File_Name), to = paste0(tmp, "/", excel_File_Name))
   setwd(tmp)
   file.rename(from = excel_File_Name, to = file_Name_Zip)
   unzip(zipfile = file_Name_Zip)
   setwd(tmp <- paste0(tmp,"/xl/drawings"))
   drawing_files <- list.files(pattern = "\\.xml")
   nb_Drawing_Files <- length(drawing_files)
   vector_Text <- character(nb_Drawing_Files)
   
   for(i in 1 : nb_Drawing_Files)
   {
      xml_Text <- read_xml(drawing_files[i])
      text <- xml_text(xml_Text, trim = TRUE)
      text <- str_replace_all(text, pattern = "\\d{12,200}", replacement = "")
      vector_Text[i] <- text
   }
   
   vector_Text <- vector_Text[str_detect(vector_Text, "[:alpha:]")]
   
   return(vector_Text)
}

#Tach phan textbox thanh table-------------------------------------------------------
for(y in 2020:2021)
{
   time_start_2 = Sys.time()
   for (e in 1:12)
   {
      # e=1
      # y   = 2019
      y_1  = str_sub(y, start = 3L); y_1
      dir  = wd
      if (nchar(e) == 1)
      {
         file_input   = paste0("0", e, ".", y)
         tb_input     = paste0("0", e, ".", y_1,".TB")
      } else {
         file_input   = paste0(e, ".", y)
         tb_input     = paste0( e, ".", y_1,".TB")
      }
      
      # if(file_input %in% c(
      #    "04.2022",
      #    "05.2022",
      #    "06.2022",
      #    "07.2022",
      #    "08.2022",
      #    "09.2022",
      #    "10.2022",
      #    "11.2022",
      #    "12.2022"))
      # {
      #    break
      # }
      
      excel_File_Name = paste0(file_input, ".xls")
      TB_File_Name    = paste0(tb_input, ".xls")
      
      {
         f         = extract_Textbox_From_Excel(dir, excel_File_Name)
         f1        = str_split(f, "Tên tài khoản")
         f1
         title     = data.table()
         new_title = data.table()
         for (i in 2:length(f1[[1]]))
         {
            # i=2
            cat("\014")
            print(paste("Tai khoan", f1[[1]][i]))
            # f1[[1]][i]
            stk       = str.extract(f1[[1]][i], ": ", " -")
            stk
            ttk       = str.extract(f1[[1]][i], " - ", "- S")
            ttk
            # new_stk   = str.extract(f1[[1]][i], ": ", "- S")
            stk       = cbind(stk, ttk)
            title     = rbind(title, stk)
            # new_title = rbind(new_title, new_stk)
            
         }
      }
      My.Kable(title)
      dup = sum(duplicated(title))
      
      
      # openxlsx::write.xlsx(title,paste0(dir, "/RESULT/So_cai/Title ", file_input, ".xlsx"))
      print(excel_File_Name)
      if (title$stk[1] == 111100001)
      {
         title$ttk[1] <- "Tiền Việt Nam"
      }
      
      #Nhap data Buoc 2------------------------------------------------------------------------
      input   = setDT(read_xlsx(paste0(wd, "/DATA_RAW/", excel_File_Name), 1))
      test_tb = setDT(read_excel(paste0(wd, "/TB/",y,"/", TB_File_Name),   1))
      file    = setDT(read_excel('C:/Users/KhaiPham/OneDrive - Grant Thornton (Vietnam) Limited/Quang Khai/Every_r_studio/file_test/Món ăn ở Đồng Hới.xlsx',   1))
      
      #Xu ly phan bang (header,...) Buoc 3-------------------------------------------------
      {
         setnames(input,   "Chứng từ", "Số hiệu")
         setnames(input,   "...3", "Ngày tháng")
         setnames(input,   "...6", "STT dòng")
         setnames(input,   "Số tiền", "Nợ")
         setnames(input,   "...9", "Có")
         setnames(input,   "Số hiệu\r\nTK đối ứng", "Số hiệu TK đối ứng")
         setnames(input,   "Ngày tháng\r\nghi sổ", "Ngày tháng ghi sổ")
         setnames(test_tb, "SỐ PHÁT SINH TRONG KỲ", "Nợ")
         setnames(test_tb, "...6", "Có")
         test_tb = test_tb[,.(`SỐ HIỆU TÀI KHOẢN`,`TÊN TÀI KHOẢN KẾ TOÁN`,Nợ, Có)]
         str(input$`Ngày tháng ghi sổ`)
         
         date1 = as.data.table(as.Date(as.numeric(input$`Ngày tháng ghi sổ`), origin = "1899-12-30"))
         setnames(date1, "V1", "Ngày tháng ghi sổ 2")
         date2 = as.data.table(as.Date(as.numeric(input$`Ngày tháng`), origin = "1899-12-30"))
         setnames(date2, "V1", "Ngày tháng 2")
      }
      
      z = as.data.table(input)
      # head(z)
      z = cbind(date1, z)
      # head(z)
      z = cbind(date2, z)
      # head(z)
      z = z[!is.na(`Diễn giải`)][, `Ngày tháng ghi sổ` := NULL][, `Ngày tháng`:= NULL]
      z1 = z[1]
      
      z = rbind(z, z1)
      summary(is.na(z))
      h = z[`STT dòng` == "F"]
      nrow(h)
      
      #Kiem tra TB -----------------------------------------------------------------------
      {
         
         ## Nhap TB (Làm ở trên luôn) -------------------------------------------------------
         
         ##Tách z thành sub_z ---------------------------------------------------------------
         sub_z   = as.data.frame(z[, .(`Số hiệu TK đối ứng`, Nợ, Có)]
                                 [, ":="(Nợ  = as.numeric(Nợ),
                                         Có  = as.numeric(Có))][order(`Số hiệu TK đối ứng`)])
         sub_z    = as.data.table(aggregate(. ~ `Số hiệu TK đối ứng`, sub_z, sum))
         test_tb = test_tb[!is.na(`SỐ HIỆU TÀI KHOẢN`)][order(`SỐ HIỆU TÀI KHOẢN`)]
         ##Mapping theo STK------------------------------------------------------------------
         test_tb = transform(test_tb, Nợ_sub_z = sub_z$Nợ[match(test_tb$`SỐ HIỆU TÀI KHOẢN`, sub_z$`Số hiệu TK đối ứng`)])
         test_tb = transform(test_tb, Có_sub_z = sub_z$Có[match(test_tb$`SỐ HIỆU TÀI KHOẢN`, sub_z$`Số hiệu TK đối ứng`)])
         test_tb = test_tb[, ":="(Nợ  = as.numeric(Nợ),
                                  Có  = as.numeric(Có))]
         test_tb = test_tb[, ":="(Chênh_lệch_nợ  = Nợ -Có_sub_z,
                                  Chênh_lệch_có  = Có -Nợ_sub_z)]
         test_tb[is.na(test_tb)] <- 0
         tb_sum_no  = sum(test_tb$Chênh_lệch_nợ)
         tb_sum_co  = sum(test_tb$Chênh_lệch_có)
         #TEST------------------------------------------------------------------------------
         if (tb_sum_co != 0)
         {
            print("Tổng Có sai với TB")
            print(paste0("file z: ", excel_File_Name))
            beep(sound = 9, expr = NULL)
            next
         } 
         
         if (tb_sum_no != 0)
         {
            print("Tổng Nợ sai với TB")
            print(paste0("file z: ", excel_File_Name))
            beep(sound = 9, expr = NULL)
            next
         }
      }
      
      
      
      if (nrow(h) == nrow(title) + 1)
      {
         setnames(z, "Ngày tháng ghi sổ 2", "Ngày tháng ghi sổ")
         setnames(z, "Ngày tháng 2", "Ngày tháng")
         My.Kable(z)
         # z = z[,Date:= as.POSIXct.numeric(as.numeric(`Ngày tháng ghi sổ`), origin = "01/01/2020")]; head(z)
         
         list_all = list()
         tab_all  = data.table()
         tab_sub  = data.table()
         j = 1
         time_start = Sys.time()
         for (i in 2:nrow(z))
         {
            # i = 457114
            cat("\014")
            print(paste("i = ", i))
            print(paste("j = ", j))
            print(paste("status i = ", z$`Diễn giải`[i]))
            print(paste("status j = ", title[j]))
            print(paste0("file = ", file_input))
            
            if (is.na(z$`STT dòng`[i]))
            {
               tit     = cbind(z[i], title[j])
               tab_sub = rbind(tab_sub, tit)
            } else {
               tit     = cbind(z[i], title[j])
               tab_sub = rbind(tab_sub, tit)
               tab_all = rbind(tab_all, tab_sub)
              
               tab_sub = data.table()
               j       = j + 1
            }
         }
         My.Kable(tab_all)
         
         print(paste0("overtime: ", format(Sys.time() - time_start, digits = 6)), quote = F)
         tab_all = tab_all[!is.na(`Số hiệu TK đối ứng`)][!is.na(`Ngày tháng`)]
         tab_all = tab_all[,":="(Nợ = as.numeric(Nợ),
                                 Có = as.numeric(Có))]
         {
         ## Nhap TB (Làm ở trên luôn) -------------------------------------------------------
         
         ##Tách z thành sub_z ---------------------------------------------------------------
         sub_all   = as.data.frame(tab_all[, .(`Số hiệu TK đối ứng`, Nợ, Có)]
                                 [, ":="(Nợ  = as.numeric(Nợ),
                                         Có  = as.numeric(Có))][order(`Số hiệu TK đối ứng`)])
         sub_all    = as.data.table(aggregate(. ~ `Số hiệu TK đối ứng`, sub_all, sum))
         ##Mapping theo STK------------------------------------------------------------------
         test_tb = transform(test_tb, Nợ_sub_all = sub_all$Nợ[match(test_tb$`SỐ HIỆU TÀI KHOẢN`, sub_all$`Số hiệu TK đối ứng`)])
         test_tb = transform(test_tb, Có_sub_all = sub_all$Có[match(test_tb$`SỐ HIỆU TÀI KHOẢN`, sub_all$`Số hiệu TK đối ứng`)])
         test_tb = test_tb[, ":="(Nợ  = as.numeric(Nợ),
                                  Có  = as.numeric(Có))]
         test_tb = test_tb[, ":="(Chênh_lệch_nợ  = Nợ - Có_sub_all,
                                  Chênh_lệch_có  = Có - Nợ_sub_all)]
         test_tb[is.na(test_tb)] <- 0
         tb_sum_no  = sum(test_tb$Chênh_lệch_nợ)
         tb_sum_co  = sum(test_tb$Chênh_lệch_có)
         #TEST------------------------------------------------------------------------------
         if (tb_sum_co != 0)
         {
            print("Tổng Có sai với TB")
            print(paste0("file all: ", excel_File_Name))
            beep(sound = 9, expr = NULL)
            break
         }
         
         if (tb_sum_no != 0)
         {
            print("Tổng Nợ sai với TB")
            print(paste0("file all: ", excel_File_Name))
            beep(sound = 9, expr = NULL)
            break
         }
      }
         openxlsx::write.xlsx(tab_all,
                              paste0(dir, "/RESULT/So_cai/", file_input, ".xlsx"),
                              overwrite = T)
         print(paste0("Đã lưu file ", file_input))
         Sys.sleep(1)
         beep(sound = 2, expr = NULL)
         
         
         
      } else {
         print("Số F khác số title")
         beep(sound = 9, expr = NULL)
      }
   }
   print(paste0("overtime: ", format(Sys.time() - time_start_2, digits = 6)), quote = F)
}
beep(sound = 8, expr = NULL)



#Kiem dinh file-----------------------------------------------------------------
{
i                = 1
y                = 2021
test_file_result = setDT(read_xlsx(paste0(wd, "/RESULT/So_cai/0", i, ".", y, ".xlsx")))
test_file_raw    = setDT(read_xlsx(paste0(wd, "/DATA_RAW/0", i, ".", y, ".xls")))
My.Kable(test_file)
test_title_stk   = as.data.table(unique(test_file_result$stk))
work_complete()
}


#Nhap va xuat thanh tung tai khoan----------------------------------------------
table_acc   = data.table()
for (j in 1:9)
{
   time_start_2 = Sys.time()
   for (i in 1:12)
   {
      time_start = Sys.time()
      for (y in 2019:2022)
      {
         if(j==3)
         {
            print("gà quá không chạy được")
            Sys.sleep(3)
            next
         }
         # j = 1
         # i = 10
         # y = 2021
         cat("\014")
         print(paste("Tài khoản đầu ", j))
         print(paste("Tháng ", i))
         print(paste("Năm ", y))
         if (nchar(i) == 1)
         {
            file_input   = paste0("0", i, ".", y)
         } else {
            file_input   = paste0(i, ".", y)
         }
         
         if(file_input == c("04.2022","05.2022","06.2022","07.2022","08.2022","09.2022",
                            "10.2022","11.2022","12.2022"))
         {
            next
         }
            
         
         if (nchar(i) == 1)
         {
            table_in    = setDT(read_xlsx(paste0(
               wd, "/RESULT/So_cai/0", i, ".", y, ".xlsx"
            )))
         } else {
            table_in   = setDT(read_xlsx(paste0(
               wd, "/RESULT/So_cai/", i, ".", y, ".xlsx"
            )))
         }
         # names(table_year) <- names(paste("table_",y))
         setnames(table_in, "Ngày tháng...1", "Ngày tháng")
         table_in    = table_in[, `Ngày tháng...4` := NULL]
         # table_in    = table_in[, ":="(Nợ  = as.integer(Nợ),
         #                               Có  = as.integer(Có),
         #                               stk = as.integer(stk))]
         head(table_in, n = 12)
         
         sub = setDT(subset(table_in, substr(stk, 1, 1) == j))
         sub = sub[, ":="(
            Tháng = month(`Ngày tháng`),
            Năm   = year(`Ngày tháng`),
            Nợ    = as.integer(Nợ),
            Có    = as.integer(Có))]
         head(sub, n = 10)
         if (nrow(sub) < 10)
         {
            print(paste0("WARNING file tháng ", i, " năm ", y))
            miss = paste0("WARNING file tai khoan dau ", j, "tháng ", i, " năm ", y)
            write.table(miss, file = paste0(wd,"/RESULT/Tai_khoan/missing_data_",i,"_",y,".txt"),sep  = "")
         }
         
         table_acc = rbind(table_acc,sub)
      }
      cat("\014")
      print(paste0("xong file tháng ", i))
      print(paste0("Thời gian chạy tháng ",i,": ", format(Sys.time() - time_start, digits = 6)), quote = F)
      Sys.sleep(2)
      beep(sound = 2, expr = NULL)
   }
   openxlsx::write.xlsx(table_acc, paste0(wd, "/RESULT/Tai_khoan/Tài khoản đầu ", j, ".xlsx"),overwrite = T)
   print(paste0("Đã lưu file tài khoản đầu ", j))
   print(paste0("Thời gian chạy tài khoản đầu ",j,": ", format(Sys.time() - time_start_2, digits = 6)), quote = F)
   beep(sound = 8, expr = NULL)
   table_acc   = data.table()
}
work_complete()



#Tach theo tung tai khoan------------------------------------------------------------
# title_all = table_allyear[, .(stk, ttk)]

# for (j in 1:9)
# {
#    cat("\014")
#    print(paste("Tài khoản đầu", j))
#    sub = setDT(subset(table_allyear, substr(stk, 1, 1) == j))
#    sub = sub[, ":="(Tháng = month(`Ngày tháng`),
#                     Năm   = year(`Ngày tháng`))]
#    head(sub, n = 10)
#    openxlsx::write.xlsx(sub,
#                         paste0(wd, "/RESULT/Tai_khoan/Tài khoản đầu ", j, ".xlsx"),
#                         overwrite = T)
# }





test_tb = setDT(read_excel(paste0(wd, "/TB/",y,"/", TB_File_Name)))
test_tb[,1]

#THE END------------------------------------------------------------------------