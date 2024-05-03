#### READ SQA ATTAINMENT PUBLICATION TABLES ####

# load packages ----
library(openxlsx)
library(dplyr)
library(purrr)
library(httr)
library(stringr)

# source functions ----
source("R/extract_publication_data.R")

# set path to excel workbook and collect data----
extract_publication_data(
  file_source = "web",
  file_name = "attainment-statistics-december-2023.xlsx",
  custom_path = NA_character_
  )