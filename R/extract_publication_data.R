#' Extract data from SQA publications
#'
#' @param file_source This is a file path or link created from [set_file_path]
#' function. Only functional for publications from 2022 onwards. 
#'
#' @returns The names for each sheets as per the contents page, the data tables
#' and the notes page. Note that shorthand has been recoded to -1 for supressed 
#' rows, -2 for not applicable and -3 for low. Dependent on the publication,
#' percentage values may be multiplied by 100
#' 

extract_publication_data <- function(file_source){
  # get the number of sheets in file ----
  sheets <- openxlsx::read.xlsx(
    xlsxFile = file_source,
    colNames = FALSE,
    sheet = 1,
    startRow = 3
    ) |> 
    dplyr::rename(sheets = X1) |> 
    dplyr::filter(stringr::str_detect(sheets, "Table"))
  
  # extract the notes sheet ----
  notes <- openxlsx::read.xlsx(
    xlsxFile = file_source,
    sheet = (nrow(sheets) + 2),
    startRow = 2
    )
  
  # check correct number of sheets
  nrow(sheets)
  
  # set start row ----
  start_row <- dplyr::if_else(
    openxlsx::read.xlsx(
      xlsxFile = file_source,
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
      xlsxFile = file_source,
      sheet = .x,
      startRow = start_row)
    )
  
  # check correct number of tables
  length(tables_raw) == nrow(sheets)
  
  percent_flag <- purrr::map(
    1:length(tables_raw),
    ~tables_raw[[.x]] |> 
      dplyr::summarise_all(~any(grepl("%", .))) |> 
      dplyr::select(dplyr::where(~. == TRUE)) |> 
      colnames()
    )
  
  # replace shorthand and convert to number ----
  ## replace suppressed values with -1, not applicable with -2 and low with -3
  ## if percentage symbol is present, remove and values should be divided by 100
  tables <- purrr::map(
    1:length(tables_raw),
    ~tables_raw[[.x]] |> 
      dplyr::mutate(
        dplyr::across(
          dplyr::everything(), 
          ~stringr::str_replace_all(.x, "\\[c\\]", "-1") |> 
            stringr::str_replace_all("\\[z\\]", "-2") |> 
            stringr::str_replace_all("\\[low\\]", "-3") |> 
            stringr::str_remove_all("%") |> 
            stringr::str_remove_all(",")
        ),
        dplyr::across(
          !dplyr::matches("Subject|Level|Qualification"),
          ~as.numeric(.x)),
        dplyr::across(
          dplyr::matches(percent_flag[[.x]]),
          ~ ifelse(. < 0, ., . / 100)
        )
      )
    ) |> 
    setNames(
      sheets |> 
        dplyr::mutate(
          sheets = sheets |> 
            stringr::str_extract_all("Table \\d*") |> 
            stringr::str_replace_all(" ", "")
        ) |>
        dplyr::pull(sheets)
      )
  
  return(list(sheets = sheets, tables = tables, notes = notes))
}

