#### READ SQA ATTAINMENT PUBLICATION TABLES ####

# load packages ----
library(openxlsx)
library(dplyr)
library(purrr)
library(httr)
library(stringr)

# source functions ----
source("R/set_file_path.R")
source("R/extract_publication_data.R")

# set path to excel workbook ----
file_path <- set_file_path(
  file_source = "web",
  file_name = "attainment-statistics-december-2023.xlsx"
  )

# collect data ----
publication_output <- extract_publication_data(file_path)

# split data ----
# Sheet titles
sheets <- publication_output$sheets

# Tables
tables <- publication_output$tables

# Notes page
notes <- publication_output$notes
