## Title:: Incorporation of Doubtful and fully verified registrations: 
## Country :: NIGERIA
## Author:: Chinedu
## Date:: October-2023


pacman::p_load(openxlsx, tidyverse, readxl,lubridate, magrittr, rlist , janitor, 
               Microsoft365R, blastula, jsonlite, glue, bizdays, httr , DBI, RMySQL, arrow) 


options(scipen = 999)

if(dir.exists("D:/Burn Manufacturing/"))  { ##Sys.info()[["user"]]=="Administrator") {
  
  server_user_path <- paste0("D:/Burn Manufacturing/")
  # user_id = "root"
  
} else {
  
  server_user_path<- paste0("C:/Users/", Sys.info()[["user"]], "/Burn Manufacturing/")
  # user_id = "enjiraini"
  
}


country_name = "NIGERIA"

country_code <- "234" ##NG

# Defining my calendar for use --------------------------------------------

# calendar <- create.calendar("BurnCalendar", weekdays=c("saturday","sunday"))
# 
# currentdate  <-  ymd(today()); currentdate
# 
# yesterday  <-  add.bizdays(Sys.Date(),-1,'BurnCalendar'); yesterday
# 


# Function to convert Excel dates to date objects -------------------------

to_date_excel <- function(dt){
  
  as.Date(as.numeric(dt), origin = "1899-12-30")
  
}


# Function to convert characters to date in R -----------------------------

to_date_r <- function(dt){
  
  as.Date(as.numeric(dt), origin = "1970-01-01")
}


# Function to Convert characters in different date formats to date --------

to_date_different_orders <- function(dt){
  
  as.Date(parse_date_time(dt, orders = c("ymd", "dmy", "mdy")))
}


# Function To Convert Date Containing Both EPOCH and Normal Date ------------------------------------------------------

convert_to_date <- function(dt){
  
  if(is.na(dt)){
    
    NA_Date_
    
  }else if(str_detect(dt, "-")){
    
    as.Date(dt)
    
  }else{
    
    to_date_excel(dt)
    
  }
  
}


# Function to trim a nd uppercase character variables ---------------------

trim_upper <- function(data) {
  data %<>%  clean_names() %>%
    mutate_if(is.character, str_to_upper) %>%  
    mutate_if(is.character, str_trim)
}

# Function to count NA ----------------------------------------------------

count_na_func <- function(x) sum((is.na(x)|x=="")) 

# Store email password ---------------------------------------------------------------------------------------------------------

Sys.setenv("SMTP_PASSWORD"="Burn@twentytwenty01") #Setting global env for the password.


# Warehouse connection ---------------------------------------------------
dsn = "BI_SERVER"
user_id = "root"
password = "Burn@2020$"

# channel <- RODBC::odbcConnect(dsn = dsn, uid = user_id,  pwd = password)
con <- dbConnect(MySQL(), dbname = "copkbi", host= "192.168.1.27",   user=user_id, password=password) ##, dbname="copkbi"

dbListTables(con)
# Pull scan out form warehouse --------------------------------------------------
# Get the current date
current_date <- Sys.Date()

# Get the first date of the current month
first_date_of_month <- as.Date(format(current_date, "%Y-%m-01"))

# Update the reg_surveys_query with the WHERE condition
reg_surveys_query <- paste0("SELECT * FROM copkbi.registrations_surveys WHERE country = '", country_name, "' AND submission_time BETWEEN '", first_date_of_month, "' AND '", current_date, "';")
reg_surveys <-  dbGetQuery(con,reg_surveys_query ) %>% mutate_all(as.character)


reg_surveys %<>% mutate(new_id = paste0(main_contact, "-", sn_mod), 
                        across(everything(),.fns = ~str_replace_all(.,"^NA,","")),
                        across(everything(),.fns = ~str_replace_all(.,"^NA$", as.character(NA))))


# Load the readr package for write_csv function
library(readr)
View(reg_surveys)


# Update the output folder to the specified path
output_folder <- "C:/Users/basil.chinedu/Documents/Doubtfully and fully verified"

# Create a file name with the current date
current_date_str <- format(Sys.Date(), format = "%Y-%m-%d")
file_name <- paste("Doubtful and Fully Verified Registrations_", current_date_str, ".xlsx", sep = "")

# Create a new Excel workbook
output_path <- file.path(output_folder, file_name)
wb <- createWorkbook()

# Write data to the first sheet
addWorksheet(wb, "REGISTRATIONS")
writeData(wb, sheet = 1, x = reg_surveys)

doubtful_verified_summary <- reg_surveys %>%
  filter(verified_cc %in% c("DOUBTFUL VERIFIED", "FULLY VERIFIED")) %>%
  group_by(verified_cc) %>%
  summarize(Total = n())


# Calculate the total of "DOUBTFUL VERIFIED" and "FULLY VERIFIED"
total_registrations <- nrow(reg_surveys)

# Count occurrences of "FULLY VERIFIED" in the verified_cc column
total_fullyverified <- sum(reg_surveys$verified_cc == "FULLY VERIFIED", na.rm = TRUE)

# Print the total count
print(total_fullyverified)

# Calculate the total of "DOUBTFUL VERIFIED" and "FULLY VERIFIED"
total_doubtfulverified <- sum(reg_surveys$verified_cc == "DOUBTFUL VERIFIED", na.rm = TRUE)

# Print the total count
print(total_doubtfulverified)

# Calculate the total of "DOUBTFUL VERIFIED" and "FULLY VERIFIED"
total_doubtful_verified <- sum(doubtful_verified_summary$Total)

# Create a summary row for the total
total_row <- data.frame(verified_cc = "TOTAL", Total = total_doubtful_verified)

# Calculate the percentage of the total and format as whole numbers with "%" symbol
doubtful_verified_summary <- doubtful_verified_summary %>%
  bind_rows(total_row) %>%
  mutate(Percentage = paste0(round((Total / total_doubtful_verified) * 100), "%"))

# Create a second sheet for the summary
addWorksheet(wb, "SUMMARY")
writeData(wb, sheet = 2, x = doubtful_verified_summary)

# Save the Excel file
saveWorkbook(wb, file = output_path)

cat("Data has been saved to:", output_path, "\n")



# ... Email Sending

library(blastula)

# Define the sender and receiver email addresses
sender_email <- "business.intelligence@burnmfg.com"
##direct_receiver <- "irene.obinikpo@burnmfg.com"
##copied_recipients <- c("etulan.ikpoki@burnmfg.com", "francis.gichere@burnmfg.com", "eric.njiraini@burnmfg.com", "antony.rono@burnmfg.com", "chidi.ohaji@burnmfg.com", "ian.juma@burnmfg.com", "basil.chinedu@burnmfg.com")
direct_receiver <- "basil.chinedu@burnmfg.com"
copied_recipients <- c("basil.chinedu@burnmfg.com")


# Create a file name with the current date
current_date_str <- format(Sys.Date(), format = "%Y-%m-%d")
file_name <- paste("Doubtful_and_Fully_Verified_Registrations_", current_date_str, ".xlsx", sep = "")

# Rename the output file
new_output_path <- file.path(dirname(output_path), file_name)
file.rename(output_path, new_output_path)


# Create an email
email <- compose_email(
  body = md(paste(
    "Hi Team,\n\nAbove is an attachment showing all the registrations this month and the breakdown of the doubtful and fully verified registrations. To enhance verification accuracy for the doubtful registrations, kindly encourage the call center team to request accurate stove serial numbers directly from customers for the affected cases.\n",
    "\nSummary Key Data Insights:",
    "\n- Total Number of Registrations:", total_registrations,
    "\n- Total Fully Verified:", total_fullyverified,
    "\n- Total Doubtfully Verified:", total_doubtfulverified,
    "\nCheers,\nBasil"
  )),
  footer = md("Sent via the Blastula - R Package")
)


# Attach the renamed Excel file to the email
email <- email %>%
  add_attachment(file = new_output_path)

# Send the email with direct receiver and copied recipients
email %>%
  smtp_send(
    to = direct_receiver,
    cc = copied_recipients,
    from = sender_email,
    subject = "Doubtful and Fully Verified Registrations",
    credentials = creds_envvar(
      user = sender_email,
      pass_envvar = "SMTP_PASSWORD",
      provider = "office365"
    )
  )

cat("Email sent successfully.\n")






