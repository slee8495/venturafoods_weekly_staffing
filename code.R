library(tidyverse)
library(lubridate)
library(readxl)

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
  writexl::write_xlsx("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 24/Weekly Staffing/2024/01.08.2024/weekly.xlsx")


########## weekly_hires_terms

weekly_hires_terms %>% 
  janitor::clean_names() %>% 
  data.frame()
