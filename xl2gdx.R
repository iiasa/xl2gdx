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
# gdxrrw: https://www.gams.com/latest/docs/T_GDXRRW.html
# tidyverse R package collection: https://www.tidyverse.org/
#
# BEWARE, on Windows installing the gdxrrw source package will not work unless
# you have a compiler installed, install a binary package instead. Binary
# packages are provided for specific operating systems and R versions,
# carefully select the appropriate package for download.
# This script uses the GAMS CSV2GDX and GDXMERGE binaries to help peform
# the conversion. These binaries are located in the GAMS system directory.
# The GAMS system directory should either be part of your PATH environment
# variable, or it can be specified via the sysdir= option.
#
# To locate the GDX libraries in the GAMS system directory, the path specified
# via the sysdir option is used if provided. Otherwise, the R_GAMS_SYSDIR
# environment variable is used if set. Otherwise the GDX libraries are loaded
# via the system-specific library search mechanism (e.g. the PATH on Windows or
# LD_LIBRARY_PATH on Linux). The GDX libraries are used via gdxrrw to write
# the output GDX.
#
# BEWARE, to guarantee that the written GDX files will load into the GAMS
# version you are using, make sure that the GAMS system directory from which
# the GDX libraries are loaded is not that of a newer GAMS version: the GDX
# format can change between GAMS versions such that older GAMS versions cannot
# load the new format,
#
# BEWARE, special characters in the Excel file are projected onto their
# closest ASCII look-alikes so as to avoid later character encoding issues.
# By limiting to ASCII, locale and platform dependencies are avoided.
# This is UNLIKE GDXRRW which stores special characters to the GDX in
# what appears to be a locale-dependent encoding.
#
# Author: Albert Brouwer

options(tidyverse.quiet=TRUE)
library(gdxrrw)
library(tidyverse)
library(readxl) # is installed when you install tidyverse
library(stringi) # is installed when you install tidyverse

# ---- Parse arguments and options ----

args <- commandArgs(trailingOnly=TRUE)
USAGE <- str_c("Usage:",
              "Rscript xl2gdx.R <Excel file> [options] [@<options file>]",
              "Options:",
              "output=<GDX file> (if omitted, output to <Excel file> but with a .gdx extension)",
              "sysdir=<GAMS system directory> (pass %gams.sysdir%)",
              "rng='<sheet>!<start_colrow>:<stop_colrow>'",
              "par=<parameter to write>",
              "cdim=<number of column dimensions>",
              "rdim=<number of row dimensions>",
              "index='<sheet>!<start_colrow>'",
              sep="\n")

VALID_OPTIONS <- c("output", "sysdir", "rng", "par", "cdim", "rdim", "index")

#TODO: remove test code below
#setwd(str_c(dirname(rstudioapi::getActiveDocumentContext()$path), "/test1"))
setwd(str_c(dirname(rstudioapi::getActiveDocumentContext()$path), "/test2"))
#args = c("test1.xls", "output=test1.gdx", "par=para", "rng=toUse!c4:f39", "rdim=1", "cdim=1")
args = c("test2.xlsx", "output=test2.gdx", "par=para", "rng=CommodityBalancesCrops1!a1:bb65501", "rdim=7", "cdim=1", "sysdir=C:\\GAMS\\win64\\27.1")
#print(args)
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

# Use given GAMS system directory to load the GDX libraries for gdxrrw
sys_dir <- options["sysdir"]
if (!is.na(sys_dir)) {
  if (!igdx(gamsSysDir=sys_dir, silent=TRUE)) {
    stop(str_glue("Cannot load GDX libraries from provided sysdir {sys_dir}"))
  }
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
rdim <- as.integer(options["rdim"])
if (is.na(rdim)) {stop("rdim option value must be an integer")}
if (is.na(options("cdim"))) {
  stop("Missing cdim option!")
}
cdim <- as.integer(options["cdim"])
if (is.na(cdim)) {stop("cdim option value must be an integer")}
if (cdim != 1) {stop("cdim != 1 not yet supported!")}

# ---- Convert Excel content to GDX ----

# Read Excel subset as a tibble, yields UTF-8 strings in case of special characters
tib <- read_excel(excel_file, range=range)

# Check whether column names are valid
col_names <- colnames(tib)
if (typeof(col_names) != "character") {
  stop("Extracted column names are not character strings!")
}
if (any(Encoding(col_names) == "UTF-8")) {
  stop(str_c("Special characters in column names not supported!: ", str_c(col_names[Encoding(col_names) == "UTF-8"], collapse=", "), collapse=""))
}

# Project latin special characters in non-value columns to ASCII.
# Unlike iconv(), stri_trans_general() yields the same results independent of locale and OS.
for (r in 1:rdim) {
  if (typeof(tib[[r]]) == "character") {
    uniq <- unique(tib[[r]])
    if (any(Encoding(uniq) == "UTF-8")) {
      uniq_proj <- stri_trans_general(uniq, "latin-ascii")
      if (any(Encoding(uniq_proj) == "UTF-8")) {
        # Non-latin special characters are present that can not be projected.
        stop(str_c("Cannot project special characters: ", str_c(unipro[Encoding(uniq_proj) == "UTF-8"], collapse=", "), collapse=""))
      }
      warning(str_c("Special characters projected to ASCII look-alikes: ", str_c(str_c(uniq[Encoding(uniq) == "UTF-8"], uniq_proj[Encoding(uniq) == "UTF-8"], sep=" -> "), collapse=", "), collapse=""))
      proj <- stri_trans_general(tib[[r]], "latin-ascii")
      tib[[r]] <- proj
    }
  }
}

# Factor non-value columns
for (r in 1:rdim) {
  tib[[r]] <- factor(tib[[r]])
}

# Gather value columns
g <- tib %>%
     gather(col_names[(rdim+1):length(col_names)], key="type", value="value") %>%
     filter(!is.na(value))

# Factor type column
g[["type"]] <- factor(g[["type"]])

# Write to GDX
attr(g, "symName") <- options["par"]
attr(g, "domains") <- col_names[1:rdim]
g
wgdx.lst(gdx_file, list(g))
