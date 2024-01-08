library(tidyverse)
library(lubridate)
library(readxl)
library(fs)

# Define the paths
old_dir <- "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/2024/01.08.2024"
new_dir <- "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/2024/01.08.2024"
file_name <- "Recap of Weekly Site Staffing Updates.xlsx"

# Create the new directory
dir.create(new_dir)

# Copy the file
file.copy(file.path(old_dir, file_name), file.path(new_dir, file_name))

# Read in data
weekly_hiring_ms_data <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/2024/01.08.2024/Weekly Site Staffing Update 01_02_2024_ Updated File1.xlsx",
                                           sheet = "Hiring MS DATA")

weekly_hires_terms <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/2024/01.08.2024/Weekly Site Staffing Update 01_02_2024_ Updated File1.xlsx",
                                        sheet = "Hires & Terms")

# recap_ms <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/Recap of Weekly Site Staffing Updates.xlsx",
#                     sheet = "MS Recap")

# recap_turnover <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/Recap of Weekly Site Staffing Updates.xlsx",
#                              sheet = "Turnover Recap")



readRDS("master_data_ms_rds.rds") -> master_data_ms_rds
# readRDS("master_data_turnover_rds.rds") -> master_data_turnover_rds

############################## Clean up data
########## weekly_hiring_ms_data

weekly_hiring_ms_data %>%
  janitor::clean_names() %>%
  dplyr::mutate(report_date = format(lubridate::ymd(report_date), "%m/%d/%Y")) %>% 
  dplyr::filter(!is.na(external_openings)) %>% 
  dplyr::rename(plant_name = plant) -> weekly_hiring_ms_data_cleaned

# recap_ms %>%
#   janitor::clean_names() %>%
#   mutate(report_date = lubridate::date(report_date)) -> recap_ms_cleaned

# rbind(recap_ms_cleaned, weekly_hiring_ms_data_cleaned) -> master_data_ms_rds
rbind(master_data_ms_rds, weekly_hiring_ms_data_cleaned) -> master_data_ms_rds

master_data_ms_rds %>%
  mutate(row_id = apply(., 1, paste, collapse = "")) %>%
  distinct(row_id, .keep_all = TRUE) %>%
  select(-row_id) -> master_data_ms_rds



saveRDS(master_data_ms_rds, "master_data_ms_rds.rds")

weekly_hiring_ms_data_cleaned %>% 
  dplyr::rename("Report Date" = report_date,
                "Plant #" = plant_number,
                "Plant Name" = plant_name,
                "Department" = department,
                "External Openings" = external_openings,
                "Internal Openings" = internal_openings,
                "Pending BG/DS" = pending_bg_ds,
                "Filled By Temps" = filled_by_temps) %>% 
  writexl::write_xlsx("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/2024/01.08.2024/ms recap.xlsx")


########## weekly_hires_terms

weekly_hires_terms %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::filter(dplyr::pull(., 1) %in% c("Albert Lea", 
                                         "Birmingham",
                                         "Chambersburg", 
                                         "Ft. Worth",
                                         "Ontario",
                                         "Opelousas",
                                         "Port St. Lucie",
                                         "Portland",
                                         "Salem",
                                         "St. Joseph",
                                         "Thornton",
                                         "Waukesha",
                                         "Torlake",
                                         "Edmonton",
                                         "Brantford"
                                        )) %>% 
  dplyr::mutate(across(c(1, ncol(.)), ~replace(., is.na(.), 0))) %>% 
  dplyr::select(1, ncol(.)) %>% 
  dplyr::rename(Location = names(.)[1], Value = names(.)[2]) -> weekly_hires_terms_cleaned

weekly_hires_terms_cleaned %>% tail(15) %>% rename(Terms = Value) -> terms
weekly_hires_terms_cleaned %>% head(15) %>% rename(Hires = Value) -> hires


cbind(terms, hires) %>% 
  writexl::write_xlsx("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/2024/01.08.2024/Turnover Recap.xlsx")



##### What to do next
# Open "ms recap.xlsx" & "Recap of Weekly Site Staffing Updates.xlsx"
# In "ms recap.xlsx", copy from B to the end. and open Recap of Weekly Site Staffing Updates.xlsx -> Tab: MS Recap
# Paste the copied data to the end of the table, and manually input the date in the first column from "ms recap.xlsx" file

# Go to "Turnover Recap" Tab in the master file (Recap of Weekly Site Staffing Updates.xlsx)
# Open "Turnover Recap.xlsx" file and copy & paste. 


