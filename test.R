#--------------- LOAD LIBRARIES ---------------

if (!require("pacman")) install.packages("pacman"); library(pacman)

p_load("tidyverse", "readxl", "openxlsx", "janitor", "purrr") 

########

clean_data <- function(df) {

  base_file_name <- basename(src_file_name)
  
  base_file_name <- gsub(base_file_name, pattern=".xls$", replacement="")
  
  print(base_file_name)
  
  #--------------- CLEAN DATA ---------------
  
  # Session -----
  
  df$`Session` <- as.numeric(df$`Session`)
  
  # Names -----
  
  df$`First Name` <- toupper(df$`First Name`)
  
  df$`Last Name` <- toupper(df$`Last Name`)
  
  # Internal White Spaces -----
  
  df$`First Name` <- gsub("\\s+"," ", df$`First Name`)
  
  df$`Last Name` <- gsub("\\s+"," ", df$`Last Name`)
  
  df$`Address` <- gsub("\\s+"," ", df$`Address`)
  
  # df <- df %>% 
  #   mutate(
  #     `Modified Date Parsed` = as.Date(`Modified Date`, format = "%m/%d/%Y")
  #   )
  # 
  # head(df$`Modified Date Parsed`)
  # 
  # as.Date(df$`Modified Date`, "%m/%d/%y")
  # 
  # 
  #  #df$`Modified Date` <- as.Date(`Modified Date`, format = "%m/%d/%Y")
  
  #--------------- ANALYSIS ---------------
  
  # Sort -----
  
  df <- df %>% 
    arrange(
      `LevSessionNumber`,
      `Address`,
      `Last Name`,
      `First Name`
    )
  
  # Remove duplicates -----
  
  df <- df %>% 
    distinct(
      `Session`,
      .keep_all = TRUE
    )
  
  # Create Key column ----- 
  
  # Concat info usually checked for matches
  df <- df %>% 
    mutate(
      `Key` = paste(
        `First Name`,
        `Last Name`,
        `SPCode`,
        `Date of Sale`,
        `LevSessionNumber`,
        sep =";"
      )
    )
  
  # Find Key duplicates -----
  
  df_key <- df %>% 
    get_dupes(
      `Key`
    ) %>% 
    arrange(
      `LevSessionNumber`
    ) %>% 
    mutate(
      `Dup Check` = "1"
    )
  
  # Invoice Check -----
  
  df_invoice_check <- df %>%
    group_by(
      `LevSessionNumber`,
      `Invoice`
    ) %>%
    summarise(
      `Invoice Check` = n()
    ) %>%
    filter(
      `Invoice Check` > 1
    ) %>% 
    mutate(
      `Invoice Check` = "1"
    )
  
  # Left Join -----
  
  df <- left_join(df, df_key)
  
  df <- left_join(df, df_invoice_check)
  
  # Insert blank Raction column -----
  
  df$`Raction` = "No Action"
  
  # Re-order columns -----
  
  df <- df %>%
    select(
      `Raction`,
      everything(),
      -`Key`
    )
  
  #--------------- REPORTING ---------------
  
  setwd("C:/Users/shenderson/Desktop/RAC_Projects/ER_Reports/Exports")
  
  
  # For tracking -----
  
  # Total hits
  total_hits_count <- nrow(df)
  
  print(paste0(total_hits_count, " - Total Hits"))
  
  # Total assessed
  total_assessed_count <- nrow(df_key)
  
  print(paste0(total_assessed_count, " - Total Assessed"))
  
  # Export file -----
  
  # Get row and column index -----
  
  last_row <- nrow(df)+1
  all_cols <- 1:ncol(df)
  
  # Create workbook -----
  
  wb <- createWorkbook()
  
  # Add sheets -----
  
  addWorksheet(
    wb, 
    sheetName = "AllSessionSorted"
  )
  
  # Write data -----
  
  writeData(
    wb, 
    sheet = "AllSessionSorted",
    x = df, 
    withFilter = TRUE
  )
  
  # Formatting styles -----
  
  # BgFill only used for conditional formatting styles only -> use fgFill
  
  red_style <- createStyle(
    fontColour = "#9C0006", 
    fgFill = "#FFC7CE"
  )
  
  yellow_style <- createStyle(
    fontColour = "#9C6500", 
    fgFill = "#FFEB9C"
  )
  
  green_style <- createStyle(
    fontColour = "#006100", 
    fgFill = "#C6EFCE"
  )
  
  blue_style <- createStyle(
    fontColour = "#006100", 
    fgFill = "#B4C6E7"
  )
  
  red_border <- createStyle(
    border = "bottom", 
    borderColour = "#FF0000", 
    borderStyle = "thick"
  )
  
  grey_style <- createStyle(
    fontColour = "#000000",
    bgFill = "#DBDBDB"
  )
  
  # Conditional formatting styles -----
  
  yellow_style_cf <- createStyle(
    fontColour = "#9C6500", 
    bgFill = "#FFEB9C"
  )
  
  # Style highlighting -----
  
  # Green - Customer info
  addStyle(
    wb,
    sheet = "AllSessionSorted",
    green_style,
    rows = 1:last_row,
    cols = 5:10,
    gridExpand = TRUE # Default is false
  )
  
  # Yellow - Program Code & Program Description
  addStyle(
    wb,
    sheet = "AllSessionSorted",
    yellow_style,
    rows = 1:last_row,
    cols = c(13, 20),
    gridExpand = TRUE
  )
  
  # Red - Model
  addStyle(
    wb,
    sheet = "AllSessionSorted",
    red_style,
    rows = 1:last_row,
    cols = 17,
    gridExpand = TRUE
  )
  
  # Red - Date of Sale
  addStyle(
    wb,
    sheet = "AllSessionSorted",
    blue_style,
    rows = 1:last_row,
    cols = 24,
    gridExpand = TRUE
  )
  
  # Conditional formatting rules -----
  
  # Red border
  conditionalFormatting(
    wb, 
    sheet = "AllSessionSorted", 
    cols = all_cols, 
    rows = 2:last_row, 
    type = "expression", 
    rule = '=NOT($V2=$V3)', 
    style = red_border
  )
  
  # Grey - paid out
  conditionalFormatting(
    wb, 
    sheet = "AllSessionSorted", 
    cols = all_cols, 
    rows = 2:last_row, 
    type = "expression", 
    rule = '=SEARCH("*Paid*",$W2)', 
    style = grey_style
  )
  
  # Yellow - yesterdays date modified
  conditionalFormatting(
    wb, 
    sheet = "AllSessionSorted", 
    cols = all_cols, 
    rows = 2:last_row, 
    type = "expression", 
    rule = '$Z2=TODAY()-1', 
    style = yellow_style_cf
  )
  
  
  # Built report filename -----
  
  # Add date to file name
  built_report_filename <- paste0(base_file_name, 
                                  " - CHECKED",
                                  ".xlsx")
  
}

##################################################

src_file_path <- paste0(Sys.getenv(c("USERPROFILE")),"\\Desktop\\RAC_Projects\\ER_Reports\\Imports")

#src_file_path <- ("C:\\Users\\shenderson\\Desktop\\RAC_Projects\\ER_Reports\\Imports")

# Change the client code here to pull report#######
src_file_pattern <- "^Fraud Results for (.*)xls$"

files <- list.files(path = src_file_path,
                            pattern = src_file_pattern,
                            full.names = TRUE)

#files <- list.files(path="path/to/dir", pattern="*.txt", full.names=TRUE, recursive=FALSE)

lapply(files, function(x) {
  df <- read_excel(x,
                   sheet = "AllSessionSorted",
                   guess_max = Inf)
  # apply function
  out <- clean_data(df)
    # write to file
    saveWorkbook(
      wb, 
      file = built_report_filename
    )
})