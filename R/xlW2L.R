#' Convert Execl Wide Table to Long Data Frame
#'
#' This function converts Execl Wide Table to Long Data Frame.
#' @param filePath the Excel file with extension .xlsx, .xls, .xlsm, and .csv.  If csv file is selected, other arguments are not needed.
#' @param shtNum an Excel sheet number for data table, defaults to 1.
#' @param celIStart an Excel data starting Cell Index (upper left corner), e.g., em218, eM218, EM218, or 218Em, 218em (with max.2-letter column header).
#' @param cellEnd an Excel data ending Cell Index (bottom right corner), e.g., dW543 (with max.2-letter column header).
#' @param x1cell an Excel Cell Index for cell with key header, which separates identification variables from Y-value varables, e.g., bf38 (with max.2-letter column header).
#' @param x1Name optional key header, defaults to "" will use the original Excel key name if it exists.
#' @param yName optional response variable name for y-values returned in Long Data Frame, defaults to "" will use "Y" for header of y-value column in Long Data Frame.
#' @return xlW2L(..)$x1T or xlW2L(..)[1] -- the original Excel table;  xlW2L(..)wideF or xlW2L(..)[2] -- the wide form after cleaning up;  xlW2L(..)longF or xlW2L(..)[3] -- the long data form. If the original file is .csv type, all three objects returned are the original csv table.
#' @keywords Excel
#' @export
#' @examples
#' xlW2L(filePath,,"e8","o44","i8")$longF
#' xlW2L(filePath)$longF -- if original file is .cvs type, it returns the original cvs table.


xlW2L <- function(filePath, shtNum = 1, celIStart, cellEnd, x1cell, x1Name = "", yName = "") {
  
  if (tools::file_ext(filePath) == "csv") {
    daf <- read.csv(filePath)
    xsnum <- ncol(daf)
    for (i in 1:xsnum) {
      if (inherits(daf[, i], "Date")) 
        daf[, i] <- format(daf[, i], "%D")  #if a column is Date, change to desired date format
    }
    W2L <- list(xlT = daf, wideF = daf, longF = daf)
    return(W2L)  #return(W2L) can be also turned off, -if it is on, both objects will be printed out when function is called.
    # The objects can be called with: W2L(...)$xlT or W2L(...)$wideF
  }
  
  if (yName == "") 
    yName <- "Y"
  
  cs <- cel2num::cel2num(celIStart)$c
  rs <- cel2num::cel2num(celIStart)$r
  ce <- cel2num::cel2num(cellEnd)$c
  re <- cel2num::cel2num(cellEnd)$r
  x1c <- cel2num::cel2num(x1cell)$c  #column index for x1 variable name
  
  if (cs == x1c) 
    cs <- cs + 1  # skip x1c column
  if (x1c < cs) 
    x1c <- cs  #x1 column can not before table starting
  
  x1num <- ce - x1c  # number of values of x1 variable
  xsnum <- x1c - cs  # number of other x variables
  
  daf <- xlsx::read.xlsx(filePath, shtNum, colIndex = c(cs:ce), rowIndex = c(rs:re), 
    check.names = FALSE, header = T)
  
  if (xsnum > 0) {
    for (i in 1:xsnum) {
      if (inherits(daf[, i], "Date")) 
        daf[, i] <- format(daf[, i], "%D")  #if a column is Date, change to desired date format
    }
  }
  
  wideF <- daf  #wideF for filled the empty of other x-variables
  
  ncd <- 0
  if (xsnum > 0) {
    # Fill the empty cell with value of last non-empty cell for variables
    # other than X1 and remove xVariable columns without header and 1st
    # cell value
    b <- NULL  #b vector for xVariable columns with first row cell of NA
    for (i in 1:xsnum) {
      while (length(ind <- which(is.na(wideF[, i]))) > 0 && !is.na(wideF[1, 
        i]) && names(wideF)[i] != "NA") {
        wideF[, i][ind] <- wideF[, i][ind - 1]
      }
      ifelse(!is.na(wideF[1, i]) && names(wideF)[i] != "NA", a <- 0, 
        b <- c(b, i))  #determine col#s of xVariables with first cell NA
      # must use !is.na(wideF[1,i]), because 'Error in if (is.na(wideF[1,
      # i])) { : argument is of length zero' NA for header is treated as
      # charaters in is.na(names(wideF)[i])
    }
    ncd <- length(b)  #numbers of columns deleted of ID columns
    if (ncd > 0) 
      wideF <- wideF[, -b]  #remove xVariable columns with first cell of NA and no headers
    if (xsnum == ncd) 
      wideF <- wideF[, -1]
  }
  
  wideD <- wideF  #filled normal wide data form
  
  # delete NA columns among X1 columns in wideF for longF processing
  xcd <- sum(names(wideD) == "NA")  #count column # with header 'NA', actually the x1 columns with 'NA' headers
  wideD <- wideD[, !names(wideD) == "NA"]  #x1 NA column deleted, it changed column # from the definition at the beginning
  
  if (xsnum - ncd < 1) {
    # no ID columns
    vari <- 1:(ce - cs + 1 - ncd - xcd)
    if (xsnum == ncd && xsnum != 0) 
      vari <- 1:(ce - cs - ncd - xcd)
    longF <- reshape(wideD, v.names = yName, varying = vari, direction = "long", 
      new.row.names = 1:(nrow(wideD) * length(vari)), times = names(wideD)[vari])  #to preserve original 'numeric' column names in long column
    # timevar=names(wideD)[xsnum+1-ncd], idvar=idvari,
    # drop=names(wideD)[xsnum+1-ncd], #drop colX1 column, ncd -- adjust for
    # col deleted new.row.names = 1:(nrow(wideD)*length(vari)), #1: new
    # long row numbers
    
    longF$id <- NULL  #delete id column
    names(longF)[1] <- "X"
  } else {
    # having ID columns
    vari <- (xsnum + 2 - ncd):(ce - cs + 1 - ncd - xcd)  # column list for X1 values (and y-values below )
    idvari <- names(wideD)[1:(xsnum - ncd)]
    longF <- reshape(wideD, timevar = names(wideD)[xsnum + 1 - ncd], 
      v.names = yName, varying = vari, idvar = idvari, drop = names(wideD)[xsnum + 
        1 - ncd], direction = "long", new.row.names = 1:(nrow(wideD) * 
        length(vari)), times = names(wideD)[vari])  #to preserve original 'numeric' column names in long column
    ## drop=names(wideD)[xsnum+1-ncd], #drop colX1 column, ncd -- adjust for
    ## col deleted new.row.names = 1:(nrow(wideD)*length(vari)), #1: new
    ## long row numbers
  }  #close if
  
  longF <- na.omit(longF)  #remove any row with blank cells
  if ((x1num - xcd) == 1) 
    longF[, ncol(longF) - 1] <- NULL  #if X1 has only one column, remove X1 value column
  if ((x1num - xcd) == 1 & yName == "Y") 
    names(longF)[ncol(longF)] <- names(wideD)[ncol(wideD)]  #change back to its original column name if X1 has only one column, remove X1 value column
  if (x1Name != "") 
    names(longF)[ncol(longF) - 1] <- x1Name
  
  W2L <- list(xlT = daf, wideF = wideF, longF = longF)
  return(W2L)  #return(W2L) can be also turned off, -if it is on, both objects will be printed out when function is called.
  # The objects can be called with: W2L(...)$xlT or W2L(...)$wideF
}  #close xlW2L function########### 