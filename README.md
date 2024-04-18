Read SQA Data Tables
================
Ryan MacGregor
April 2024

# Background

The Scottish Qualifications Authority (SQA) [publish official statistics
and other information](https://www.sqa.org.uk/sqa/48269.html) on a
regular basis. In line with the [Code of Practice for
Statistics](https://code.statisticsauthority.gov.uk/), a rolling
12-month [publication schedule](https://www.sqa.org.uk/sqa/48513.html)
is available, with an exact date of publication provided at least four
weeks in advance of publication. Further information is [available on
the SQA website.](https://www.sqa.org.uk/sqa/92537.html)

This code extracts the data contained within the data tables to allow
for analysis by users.

# Limitations

The code will only work for SQA statistical publications with the
updated excel layout, this will exclude publications prior to 2022.

# R Packages

The following R packages are required to be installed:

- openxlsx
- dplyr
- purrr
- httr
- stringr

# Workflow

Run each of the following sections in order:

- Load packages and source functions.
  - **Description**: Loads required R packages and functions for reading
    data tables.
- Set path to excel workbook
  - **Description**: Creates the file patch for the data tables. Can be
    set to a web link or local file path.
  - **Scripts**: Runs the R script `set_file_path.R`.
  - **Output**: Outputs a variable called `file_path` in the R
    environment.
- Collect data
  - **Description**: Takes file path created in ‘Set path to excel
    workbook’ section and extracts sheet names, data tables, and notes
    to a nested list.
  - **Scripts**: Runs the R script `extract_publication_data.R`.
  - **Output**: Outputs a nested list `publication_output` in the R
    environment.
- Split data
  - **Description**: Splits data collected in ‘Collect data’ section.
  - **Output**: Outputs a list of data called `tables` in the R
    environment, as well as data frames called `sheets` and `notes`.

# Project structure/archive

The parent directory is an RStudio project initiated as a git repo and
includes the associated admin files (`.gitignore` and
`read_sqa_publications.Rproj`) at the top level. It also contains the
following sub-directory and a taskmaster file:

## `R/`

- `set_file_path.R`
  - contains a function `set_file_path()` that creates a file path to
    the location of an SQA statistical excel publication either from the
    web or on the local device.
- `extract_publication_data.R`
  - contains a function `extract_publication_data()` that extracts the
    sheet names, data tables, and notes of an SQA statistical excel
    publication to a nested list.

## `read_published_tables.R`

Script which runs to collect data, calling other scripts in `R/`.
Contains some variables which will need updated.