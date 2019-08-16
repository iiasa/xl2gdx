# Convert Excel to GDX
#
# This can replace GDXXRW for Excel-to-GDX conversion and accepts the same
# arguments and a subset of the options that GDXXRW does, Unlike GDXXRW,
# this script Works on non-Windows platforms and does not require Office.
#
# For further information, see the GDXXRW documentation at:
# https://www.gams.com/latest/docs/T_GDXXRW.html
#
# Requirements:
# tidyverse R package collection: https://www.tidyverse.org/
# gdxrrw: https://www.gams.com/latest/docs/T_GDXRRW.html
#
# Beware, on Windows installing the gdxrrw source package will not work unless
# you have a compiler installed, install a binary package instead. Binary
# packages are provided for specific operating systems and R versions,
# carefully select the appropriate package for download.
# This script uses the GAMS CSV2GDX and GDXMERGE binaries to help peform
# the conversion. These binaries are located in the GAMS system directory.
# The GAMS system directory should either be part of your PATH environment
# variable, or it can be specified via the sysdir= option.
#
# Author: Albert Brouwer
#
# Todo:
# Add sysdir parameter and documentation for igdx with NULL default.

options(tidyverse.quiet=TRUE)
library(tidyverse)
library(readxl) # is installed when you install tidyverse
library(gdxrrw)
igdx("C:\\GAMS\\win64\\27.1")
#quit(save="no")

# ---- Parse arguments and options ----

args <- commandArgs(trailingOnly=TRUE)
USAGE <- str_c("Usage:",
              "Rscript xl2gdx.R <Excel file> [options] [@<options file>]",
              "Options:",
              "output=<GDX file> (if omitted, output to <Excel file> but with a .gdx extension)",
              "sysdir=<GAMS system directory> (pass %gams.sysdir%, if not csv2gdx and gdxmerge must be on-path)",
              "rng=<sheet>!<start_colrow>:<stop_colrow>",
              "par=<parameter to write>",
              "cdim=<number of column dimensions>",
              "rdim=<number of row dimensions>",
              "index=<sheet>!<start_colrow>",
              sep="\n")

VALID_OPTIONS <- c("output", "sysdir", "rng", "par", "cdim", "rdim", "index")

#TODo: remove test code below
setwd(str_c(dirname(rstudioapi::getActiveDocumentContext()$path), "/test1"))
args = c("test1.xls", "output=test1.gdx", "par=para", "rng=toUse!c4:f39", "rdim=1", "cdim=1")
print(args)
#quit(save="no")


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
        stop(str_glue("Unknown option: '{args[i]}'!"))
      }
    }
  }
}
if (!is.na(options_file)) {
  if (!(file.exists(options_file))) {stop(str_glue("Options file does not exist!: '@{options_file}'"))}
}

# Stick the options into a dict-like for indexing by name
options <- structure(option_matches[,3][!is.na(option_matches[,1])], names=str_to_lower(option_matches[,2][!is.na(option_matches[,1])]))

# Use given GDX output file, or set default
gdx_file <- options["output"]
if (is.na(gdx_file)) {
  gdx_file <- str_c(excel_base_path, ".gdx")
}

# Use given GAMS system directory to find csv2gdx and gdxmerge binaries, or default to on-path
csv2gdx <- "csv2gdx"
gdxmerge <- "gdxmerge"
if (!is.na(options["sysdir"])) {
  sep = ""
  if (!(str_sub(options["sysdir"], -1) %in% c("/", "\\"))) {
    sep = "/"
  }
  csv2gdx <- str_c(options["sysdir"], csv2gdx, sep=sep)
  gdxmerge <- str_c(options["sysdir"], gdxmerge, sep=sep)
}

# Check any provided range
range = NULL
if (!is.na(options["rng"])) {
  range <- options["rng"]
  if (is.na(str_match(options["rng"], "^.+[!].+[:].+$"))) {
    stop(str_glue("Invalid rng option: '{range}' format should be <sheet>!<start_colrow>:<end_colrow>!"))
  }
}

# Check rdim and cdim
if (is.na(options("rdim"))) {
  stop("Missing rdim option!")
}
if (is.na(options("cdim"))) {
  stop("Missing cdim option!")
}

# ---- Convert Excel content to GDX ----

# Read Excel subset as a tibble
tib <- read_excel(excel_file, range=range)

# Factor non-value columns
for (r in 1:options["rdim"]) {
  tib[[colnames(tib)[r]]] <- factor(tib[[colnames(tib)[r]]])
}

# Write tibble to GDX
attr(tib, "symName") <- options["par"]
attr(tib, "domains") <- colnames(tib)[1:options["rdim"]]
tib
wgdx.lst(gdx_file, list(tib))
