#' Set file path to import SQA excel statistical data
#'
#' Sets a file path to the location of an SQA statistical excel publication.
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
#' @examples set_file_path(file_source = "web", file_name = "attainment-statistics-december-2023.xlsx")

set_file_path <- function(file_source = "web",
                          file_name = NA_character_,
                          custom_path = NA_character_){
  
  ## check valid source option has been input
  stopifnot('Must be either "web" or "custom". If "custom" provide the file path' = 
              file_source %in% c("web", "custom"))
  
  ## check file name has been provided
  stopifnot('"file_name" has not been provided, please provide the file name of the excel document' = 
              is.na(file_name) == FALSE)
  
  if (file_source == "web"){
    ## source link from web
    link <- file.path("https://www.sqa.org.uk/sqa/files_ccc", file_name)
    
    ## check if file exists, if not ask to check file name
    if (httr::status_code(httr::HEAD(link)) == 200){
      link
    }else{
      stop(paste("File can't be found, please check the provided name of the excel file:",
                 paste0('"', file_name,'"')))
    }
    
  }else if (file_source == "custom"){
    # check path has been provided
    stopifnot('"custom_path" has not been provided, please provide the file path to excel location' = 
                is.na(custom_path) == FALSE)
    
    ## source link from custom path
    link <- file.path(custom_path, file_name)
    
    ## check file exists, if not ask to check file path and file name
    if (file.exists(link)){
      link
    }else{
      stop(paste("File can't be found, please check the provided file path and name of the excel file.",
                 "\nFile path:", paste0('"', custom_path,'"'),
                 "\nFile name:" , paste0('"', file_name,'"')))
    }
    
  }
}