# Converts Excel to GDX files.
#
# This can replace GDXXRW for Excel-to-GDX conversion and accepts the same
# arguments and a subset of the options that GDXXRW does, Unlike GDXXRW,
# this script Works on non-Windows platforms and does not require Office.
#
# For further information, see the GDXXRW documentation at:
# https://www.gams.com/latest/docs/T_GDXXRW.html
#
# Required packages:
# tidyverse package collection: https://www.tidyverse.org/
# gdxrrw: https://www.gams.com/latest/docs/T_GDXRRW.html

# Beware, on Windows installing the gdxrrw source package will not work unless
# you have a compiler installed, install a binary package instead. Binary
# packages are provided for specific operating systems and R versions,
# carefully select the appropriate package for download.
#
# Author: Albert Brouwer

options(tidyverse.quiet=TRUE)
library(tidyverse)
library(readxl) # is installed when you install tidyverse
library(gdxrrw)

# ---- Parse arguments and options ----

args <- commandArgs(trailingOnly=TRUE)
USAGE <- str_c("Usage:",
              "Rscript xl2gdx.R <Excel file> [output=<GDX file>] [options] [@<options file>]",
              "Options:",
              "rng=<sheet>!<start_colrow>:<stop_colrow>",
              "par=<parameter to write>",
              "cdim=<number of column dimensions>",
              "rrim=<number of row dimensions>",
              "index=<sheet>!<start_colrow>",
              sep="\n")

VALID_OPTIONS <- c("output", "o", "rng", "par", "cdim", "rdim", "index")

args = c(readxl_example("datasets.xls"), "output=foooooo.gdx")


# Display usage if needed
if (length(args) == 0) {
  stop(str_c("Missing arguments!", USAGE, sep="\n"))
}
if (args[1] == "?" || args[1] == "--help") {
  cat(USAGE)
  quit(save="no")
}

# Match option arguments
OPTION_REGEX = "^([:alpha:]+)=(.*)$"
option_matches = str_match(args, OPTION_REGEX)

# Ensure that the first argument is the Excel file
if (str_sub(args[1], 1, 1) == "@" || !is.na(option_matches[1, 1])) {
  stop("First argument must be an Excel file!")
}
excel_file = args[1]
if (!(file.exists(excel_file))) {stop(str_glue("Excel file does not exist!: '{excel_file}'"))}
excel_base_path <- str_match(excel_file, "^(.+)[.][xX][lL][sS][xX]?$")[2]
if (is.na(excel_base_path)) {stop(str_glue("Not an Excel file: absent .xls or .xlsx extension in first argument '{excel_file}'!"))}

# Determine whether an options file has been specified, and that remaining arguments are valid options
options_file <- NA
if (length(args) > 1) {
  for (i in 2:length(args)) {
    if (is.na(option_matches[i, 1])) {
      # not an option argument
      if (is.na(options_file) && str_sub(args[i], 1, 1) == "@") {
        options_file <- str_sub(args[i], 2)
        next
      }
      stop(str_glue("Invalid argument: '{args[i]}'!"))
    } else {
      # is the option argument valid?
      if (!(str_to_lower(option_matches[i, 2]) %in% VALID_OPTIONS)) {
        stop(str_glue("Invalid option: '{args[i]}'!"))
      }
    }
  }
}
if (!is.na(options_file)) {
  if (!(file.exists(options_file))) {stop(str_glue("Options file does not exist!: '@{options_file}'"))}
}

# Stick the options into a dict-like for indexing by name
options <- structure(option_matches[,3][!is.na(option_matches[,1])], names=str_to_lower(option_matches[,2][!is.na(option_matches[,1])]))

# Determine the GDX output file
gdx_file <- options["output"]
if (is.na(gdx_file)) {
  gdx_file <- options["o"]
}
if (is.na(gdx_file)) {
  gdx_file <- str_c(excel_base_path, ".gdx")
}

# ---- Convert Excel to GDX ----

var2gdx <- function(gdx, var){
  uels <- list()
  for(n in 1:(length(var$uels) - 1)){
    uels[[n]] <- list(name=var$domains[[n]], type="set", uels=list(var$uels[[n]]))
  }
  wgdx(gdx, var, uels)
}
tib <- read_excel(excel_file)
print(tbl, n=Inf)
var2gdx(gdx_file, tbl)
