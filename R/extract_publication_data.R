#' Extract data from SQA publications
#' 
#' A function to extract all tables from the accessible versions of SQA
#' statistical publications. Replaces shorthand for suppressed values with -1, 
#' not applicable with -2 and low with -3.
#'
#' Sets the file path to the location of an SQA statistical excel publication.
#' Default is to extract from the SQA website, but if the file has been 
#' downloaded, a custom file path can be provided. Checks to see if the file 
#' exists before creating the file path for further use.
#'
#' @param file_source Options are either "web" or "custom". Default value is
#' "web", which provides a direct file path to SQA publication location. If a 
#' spreadsheet has been downloaded, then a custom file path can be used with 
#' the files location using "custom" and then providing the file path.
#' 
#' @param file_name String containing the name of the excel file required.
#' 
#' @param custom_path String containing the custom file path of the excel file.
#' Required if file_source is set to "custom".
#'
#' @returns The names for each sheets as per the contents page, the data tables
#' and the notes page. Note that shorthand has been recoded to -1 for suppressed 
#' rows, -2 for not applicable and -3 for low.
#' 
#' @import dplyr stringr
#' @importFrom openxlsx read.xlsx
#' @importFrom purrr map
#' @importFrom httr status_code HEAD
#'  
#' @examples
#' extract_publication_data(
#'   file_source = "web",
#'   file_name = "attainment-statistics-december-2023.xlsx"
#'   )
#'

extract_publication_data <- function(file_source = "web",
                                     file_name = NA_character_,
                                     custom_path = NA_character_){
  
  # input validation checks ----
  ## check valid file source option has been input, if not tell user to check
  stopifnot('Must be either "web" or "custom". If "custom" provide the file path' = 
              file_source %in% c("web", "custom"))
  
  ## check valid file name has been provided, if not tell user to check
  stopifnot('"file_name" has not been provided, please provide the file name of the excel document' = 
              is.na(file_name) == FALSE)
  
  # link and file path validation checks for web source ----
  if (file_source == "web"){
    ## source link from web
    link <- file.path("https://www.sqa.org.uk/sqa/files_ccc", file_name)
    
    ## check if file exists on web
    ## if file does not exist, stop code and tell user to check file name
    if (httr::status_code(httr::HEAD(link)) == 200){
      link
    }else{
      stop(paste("File can't be found, please check the provided name of the excel file:",
                 paste0('"', file_name,'"')))
    }
    
  }
  
  # link and file path validation checks for custom source ----
  if (file_source == "custom"){
    ## check a custom file path has been provided when custom option selected
    stopifnot('"custom_path" has not been provided, please provide the file path to excel location' = 
                is.na(custom_path) == FALSE)
    
    ## source link from custom path
    link <- file.path(custom_path, file_name)
    
    ## check file exists in custom path
    ## if file does not exist, stop code and tell user to check file path and name
    if (file.exists(link)){
      link
    }else{
      stop(paste("File can't be found, please check the provided file path and name of the excel file.",
                 "\nFile path:", paste0('"', custom_path,'"'),
                 "\nFile name:" , paste0('"', file_name,'"')))
    }
  }
  
  # get the number of sheets in file ----
  sheets <- openxlsx::read.xlsx(
    xlsxFile = link,
    colNames = FALSE,
    sheet = 1,
    startRow = 3
    ) |> 
    dplyr::rename(sheets = X1) |> 
    dplyr::filter(stringr::str_detect(sheets, "Table"))
  
  # extract the notes sheet ----
  ## notes sheet is always the last sheet in publications, add two to number of
  ## tables to account for contents page
  notes <- openxlsx::read.xlsx(
    xlsxFile = link,
    sheet = (nrow(sheets) + 2),
    startRow = 2
    )
  
  # expected tables ----
  ## tell user the number of tables to be extracted and what they should be
  cat(paste("There should be",
            nrow(sheets),
            "tables extracted from this publication:\n"), 
      paste(sheets$sheets, "\n"))
  
  # set start row ----
  ## look at cell A2 in first data table sheet for specific text used for 
  ## shorthand note identify start of data table
  start_row <- dplyr::if_else(
    openxlsx::read.xlsx(
      xlsxFile = link,
      sheet = 2,
      rows = 2,
      cols = 1,
      colNames = FALSE) |>
      stringr::str_detect("Some") == TRUE, 
    3, 
    4)
  
  # read all tables into a list ----
  tables_raw <- purrr::map(
    2:(nrow(sheets)+1),
    ~openxlsx::read.xlsx(
      xlsxFile = link,
      sheet = .x,
      startRow = start_row)
    )
  
  # check correct number of tables ----
  if(length(tables_raw) == nrow(sheets)){
    ## note to user that the correct number of tables have been extracted
    cat("\nThere were", 
        length(tables_raw), 
        "tables extracted from this publication.\n")
  }else{
    ## stops code if the number of tables extracted does not match expected output
    stop("Number of tables extracted does not match expected number of tables.")
  }
  
  # replace shorthand and convert to number ----  
  ## flag any columns with a percentage symbol so these can be altered
  percent_flag <- purrr::map(
    1:length(tables_raw),
    ~tables_raw[[.x]] |> 
      dplyr::summarise_all(~any(grepl("%", .))) |> 
      dplyr::select(dplyr::where(~. == TRUE)) |> 
      colnames()
    )
  
  ## replace suppressed values with -1, not applicable with -2 and low with -3
  ## if percentage symbol is present, remove and values should be divided by 100
  tables <- purrr::map(
    1:length(tables_raw),
    ~tables_raw[[.x]] |> 
      dplyr::mutate(
        ### replace shorthand and remove commas and percentage symbols.
        dplyr::across(
          dplyr::everything(), 
          ~stringr::str_replace_all(.x, "\\[c\\]", "-1") |> 
            stringr::str_replace_all("\\[z\\]", "-2") |> 
            stringr::str_replace_all("\\[low\\]", "-3") |> 
            stringr::str_remove_all("%") |> 
            stringr::str_remove_all(",")
          ),
        ### convert character number columns to numeric values
        dplyr::across(
          !dplyr::matches(
            "Subject|Level|Qualification|Category|SIMD.Decile|Centre.Type|
            |Education.Authority|Arrangements|Component.[0-9].Name|
            |National.[4-5].Title|[Higher|Advanced.Higher].Title|
            |National.5.Grade|Higher.Grade"),
          ~as.numeric(.x)),
        ### change percentages to their decimal value
        dplyr::across(
          dplyr::matches(percent_flag[[.x]]),
          ~ifelse(. < 0, ., . / 100)
          )
        ) |>
      ### tidy column names
      dplyr::rename_with(.fn = ~stringr::str_replace_all(., "[.]", "_"))
    ) |>
    ## set name for each element of list to associated table number
    setNames(
      sheets |> 
        dplyr::mutate(
          sheets = sheets |> 
            stringr::str_extract_all("Table \\d*.\\d*") |> 
            stringr::str_replace_all(" ", "") |> 
            stringr::str_remove_all(":")
        ) |>
        dplyr::pull(sheets)
      )
  
  ## text to notify user that shorthand values have been replaced
  cat("\nShorthand has been replaced as follows:\n[c] with -1 (suppressed),\n[z] with -2 (not applicable),\n[low] with -3\n")
  
  # function returns ----
  return(list(sheets = sheets, tables = tables, notes = notes))
}
