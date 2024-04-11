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
  - **Description**: Loads requires R packages and functions required
    for reading data tables.
- Set path to excel workbook
  - **Description**: Creates the file patch for the data tables. Can be
    set to the web or local file.
  - **Scripts**: `set_file_path.R`
  - **Output**: `file_path`
- Collect data
  - **Description**: Takes file path created above and extracts sheet
    names, data tables, and notes to a nested list.
  - **Scripts**: `extract_publication_data.R`
  - **Output**: `publication_output`
- Split data
  - **Description**: Splits data collected in previous section.
  - **Output**: `sheets`, `tables`, and `notes`

# Project structure/archive

The parent directory is an RStudio project initiated as a git repo and
includes the associated admin files (`.gitignore` and
`read_sqa_publications.Rproj`) at the top level. It also contains the
following sub-directory and a taskmaster file:

## `R/`

- `set_file_path.R`
  - functions to create a file path either from the web or local device.
- `extract_publication_data.R`
  - function to extract sheet names, data tables, and notes to a nested
    list.

# `read_published_tables.R`

Script which runs to collect data, calling other scripts in `R/`.
Contains some variables which will need updated.
