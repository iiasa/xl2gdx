#!/usr/bin/env Rscript
# Convert Excel to GDX
#
# This can replace GDXXRW for Excel-to-GDX conversion and accepts the same
# arguments and a subset of the options that GDXXRW does. Unlike GDXXRW,
# this script Works on non-Windows platforms and does not require Office.
#
# For further information, see the the USAGE definition below and the GDXXRW
# documentation at https://www.gams.com/latest/docs/T_GDXXRW.html
#
# Requirements:
# - an R installation that is not too old: tests pass with R V3.5.1 and V3.6.1
# - gdxrrw R package: https://www.gams.com/latest/docs/T_GDXRRW.html
# - tidyverse R package collection: https://www.tidyverse.org/
#
# NOTE, on Windows installing the gdxrrw source package will not work unless
# you have a compiler installed, install a binary package instead. Binary
# packages are provided for specific operating systems and R versions,
# carefully select the appropriate package for download.
#
# To locate the GDX libraries in the GAMS system directory, the path specified
# via the sysdir option is used if provided. Otherwise, the R_GAMS_SYSDIR
# environment variable is used if set. Otherwise the GDX libraries are loaded
# via the system-specific library search environment variable: PATH on Windows,
# LD_LIBRARY_PATH on Linux, or DYLD_LIBRARY_PATH on macOS. The GDX libraries
# are used via gdxrrw to write the output GDX.
#
# BEWARE, to guarantee that the written GDX files will load into the GAMS
# version you are using, make sure that the GAMS system directory from which
# the GDX libraries are loaded is not that of a newer GAMS version: the GDX
# format can change between GAMS versions such that older GAMS versions cannot
# load the new format.
#
# NOTE, unlike GDXXRW, sheet names in rng attributes are case sensitive.
#
# Author: Albert Brouwer
#
# Todo:
# - support set=?
# - support ASCII projection for headers and dsets?

start_time <- Sys.time()

VERSION <- "beta"
DATE <- "2-Oct-2019"
RESHAPE <- TRUE # select wgdx.reshape (TRUE) or dplyr-based (FALSE) parameter writing
GUESS_MAX <- 200000 # rows to read for guessing column type, decrease when memory runs low, increase when guessing goes wrong
TRIM_WS <- TRUE # trim leading and trailing whitespace from Excel fields? GDXXRW does this.

options(scipen=999) # disable scientific notation
options(tidyverse.quiet=TRUE)

library(gdxrrw)
library(tidyverse)
library(cellranger) # installed when you install tidyverse
library(readxl) # installed when you install tidyverse
library(stringi) # installed when you install tidyverse

# ---- Get command line arguments, or provide test arguments when running from RStudio ----

if (Sys.getenv("RSTUDIO") == "1") {
  # Failing argument parsing test cases
  #args <- c() # no arguments, error with usage
  #args <- c("output=foo") # first argument may not be an option
  #args <- c("@options_file") # first argument may not be an options file
  #args <- c("dummy.bad") # no Excel extension for first argument
  #args <- c("does_not_exist.xls") # not-existent Excel file
  #args <- c("dummy.xls") # an xls, but no symbol.
  #args <- c("dummy.xlsx") # an xlsx, but no symbol
  #args <- c("dummy.xLsX") # an xlsx, but no symbol
  #args <- c("dummy.xlsx", "invalid") # additional non-option argument that is not an options file
  #args <- c("dummy.xlsx", "invalid", "@options_file", "@another_options_file") # additional non-option argument that is not an options file
  #args <- c("dummy.xlsx", "@options_file", "output=foo") # options file is not the last argument
  #args <- c("dummy.xlsx", "bad=option") # unknown option
  #args <- c("dummy.xlsx", "par=foo", "output=bar") # symbol before option
  #args <- c("dummy.xlsx", "output=foo", "par=bar", "sysdir=baz") # option after symbol
  #args <- c("dummy.xlsx", "cdim=1") # symbol attribute without symbol
  #args <- c("dummy.xlsx", "cdim=1", "par=foo") # attribute without preceding symbol
  #args <- c("dummy.xlsx", "par=bar") # only symbol without attributes
  #args <- c("dummy.xlsx", "dset=foo", "par=bar", "rdim=1") # first symbol without attributes
  #args <- c("dummy.xlsx", "par=foo", "rng=A1", "rng=B2") # symbol with multiple attributes of the same type
  #args <- c("dummy.xlsx", "par=foo", "rng=invalid") # symbol with an invalid rng
  #args <- c("dummy.xlsx", "par=foo", "rng=A1") # missing cdim attribute for par
  #args <- c("dummy.xlsx", "par=foo", "rng=A1", "cdim=1") # missing rdim attribute for par
  #args <- c("dummy.xlsx", "dset=foo", "rng=A1") # missing rdim attribute for dset
  #args <- c("dummy.xlsx", "set=foo", "rng=bar!A1:B2") # no end col/row allowed for a set
  #args <- c("dummy.xlsx", "par=foo", "rng=bar!A1:B2", "cdim=invalid") # non-integer cdim
  #args <- c("dummy.xlsx", "par=foo", "rng=bar!A1:B2", "rdim=invalid") # non-integer rdim
  #args <- c("dummy.xlsx", "par=foo", "rng=bar!A1:B2", "cdim=0") # too small cdim
  #args <- c("dummy.xlsx", "par=foo", "rng=bar!A1:B2", "rdim=0") # too small rdim
  #args <- c("dummy.xlsx", "par=foo", "rng=bar!A1:B2", "cdim=1", "rdim=1", "project=invalid") # invalid value for project
  #args <- c("dummy.xlsx", "dset=foo", "rng=A1", "rdim=1", "project=Y") # project only supported for par
  args <- c("dummy.xlsx", "sysdir=C:/work", "dset=foo", "rng=A1", "rdim=1") # invalid sysdir
  
  # Conversion tests
  #args <- c("test.xls",  "testdir=test1", "par=para",   "rng=toUse!c4:f39",               "cdim=1", "rdim=1")
  #args <- c("test.xlsx", "testdir=test2", "par=para",   "rng=CommodityBalancesCrops1!a1", "cdim=1", "rdim=7", "project=N") # Re-representing UTF-8 as ASCII+latin
  #args <- c("test.xlsx", "testdir=test2", "par=para",   "rng=CommodityBalancesCrops1!a1", "cdim=1", "rdim=7", "project=Y") # Projecting UTF-8 to ASCII
  #args <- c("test.xlsx", "testdir=test3", "dset=doset", "rng=TradeSTAT_LiveAnimals1!f2",            "rdim=1")
  #args <- c("test.xlsx", "testdir=test4", "par=para",   "rng=Sheet1!AV2:BA226",           "cdim=1", "rdim=2", "par=parb", "rng=Sheet1!B2:AT226", "cdim=1", "rdim=2")
  #args <- c("test.xlsx", "testdir=test5", "par=para",   "rng=A1",                         "cdim=1", "rdim=1")
  #args <- c("test.xls",  "testdir=test6", "@taskin1.txt")
  #args <- c("test.xls",  "testdir=test7", "index=Index!B4")
  #args <- c("test.xls",  "testdir=test8", "index=INDEX!B4")
  #args <- c("test.xlsx", "testdir=test9", "par=para", "rng=Sheet2!c1:d107", "cdim=1", "rdim=1")
  #args <- c("test.xls", "testdir=test10", "par=para", "rng=PriceSTAT1!a1", "cdim=1", "rdim=8")
  #args <- c("test.xls", "testdir=test11", "@taskin.txt")
  #args <- c("test.xlsx", "testdir=test12", "par=EXCRET_MONOGAST_DATA", "rng=N_excretion!A3", "cdim=2", "rdim=2")
  #args <- c("test.xls", "testdir=test13", "index=INDEX!B4")
  #args <- c("test.xls", "testdir=test14", "par=FoodBalanceSheets2", "rng=FoodBalanceSheets2!a1:aw64001", "cdim=1", "rdim=6")
  #args <- c("test.xlsx", "testdir=test15", "par=spacey", "rng=Sheet1!B2", "cdim=1", "rdim=2")
} else {
  args <- commandArgs(trailingOnly=TRUE)
}

# ---- Display usage if needed ----

USAGE <- str_c("Usage:",
              "[Rscript ]xl2gdx.R <Excel file> [options] [@<options file>] [symbols]",
              "Prefixing with Rscript is not necessary when invoking from a Linux/MacOS shell.",
              "",
              "Global options (provide these first):",
              "    output=<GDX file> (if omitted, output to <Excel file> but with a .gdx extension)",
              "    index='<sheet>!<start_colrow>'",
              "    sysdir=<GAMS system directory> (pass %gams.sysdir%)",
              "Symbol options (one or more):",
              "    dset=<name of domain set to write>",
              "    par=<name of parameter to write>",
              "    set=<name of set to write>",
              "Symbol attribute options (associated with preceeding symbol):",
              "    cdim=<number of column dimensions>",
              "    rdim=<number of row dimensions>",
              "    rng='[<sheet>!]<start_colrow>[:<end_colrow>]'",
              "    project=Y (project latin special characters to ASCII for par symbols, defaults to N)",
              sep="\n")

# No arguments? Error with usage.
if (length(args) == 0) {
  stop(str_c("No arguments!", USAGE, sep="\n"))
}

# Display usage if requested
if (args[1] == "?" || args[1] == "-help" || args[1] == "--help") {
  cat(USAGE)
  quit(save="no")
}

# ---- Define support functions ----

# Convert Excel range string to a cell_limits object.
range2cell_limits <- function(range) {
  # The expected range format is: [<sheet>!]<start_colrow>[:<end_colrow>]
  ma <- str_match(range, "^(?:([^!]+)[!])?([:alpha:]+[:digit:]+)(?:[:]([:alpha:]+[:digit:]+))?$")
  if (is.na(ma[[1]])) {
    stop(str_glue("Invalid Excel range '{range}'. Format should be [<sheet>!]<start_colrow>[:<end_colrow>]."))
  }
  if (!is.na(ma[[4]])) {
    # A range with both a start and end col/row
    cl <- as.cell_limits(range)
  } else {
    # A range without end col/row
    cl <- anchored(ma[[3]], c(NA,NA))
    if (!is.na(ma[[2]])) {
      # Sheet name provided, add it
      cl$sheet <- ma[[2]]
    }
  }
  return(cl)
}

# ---- Preliminary parse and check of command line arguments ----

# Match (keyword=value) options and get their names and values
option_matches <- str_match(args, "^([:alpha:]+)=(.*)$")
onames <- str_to_lower(option_matches[,2][!is.na(option_matches[,1])])
values <- option_matches[,3][!is.na(option_matches[,1])]

# Stick the preliminary options (without options file or index sheet) into a dictionary
if (length(onames) > 0) {
  preliminary_options <- as.list(structure(values, names=onames))
} else {
  preliminary_options <- list()
}
rm(onames)
rm(values)

# Ensure that the first argument is a not an option nor an options file
if (str_sub(args[[1]], 1, 1) == "@" || !is.na(option_matches[[1, 1]])) {
  stop("First argument must be an Excel file!")
}

# Check that the first argument has an Excel file extension
excel_file <- args[[1]]
extensionless_excel_name <- str_match(basename(excel_file), "^(.+)[.][xX][lL][sS][xX]?$")[2]
if (is.na(extensionless_excel_name)) {stop(str_glue("Not an Excel file: absent .xls or .xlsx extension in first argument '{excel_file}'!"))}

# Determine whether an options file has been specified and is the last argument
options_file <- NA
if (length(args) > 1) {
  for (i in 2:length(args)) {
    if (is.na(option_matches[[i, 1]])) {
      # not an option argument, must be the options file
      if (str_sub(args[[i]], 1, 1) == "@") {
        if (is.na(options_file)) {
          options_file <- str_sub(args[[i]], 2)
          if (i != length(args)) {
            stop(str_glue("Invalid argument: '{args[[i]]}'! An options file must be the last argument."))
          }
        }
      } else {
        stop(str_glue("Invalid argument: '{args[[i]]}'!"))
      }
    }
  }
}
rm(option_matches)

# Change current directory for testing
if ("testdir" %in% names(preliminary_options)) {
  setwd(str_c(dirname(rstudioapi::getActiveDocumentContext()$path), "/", preliminary_options$testdir))
}
if ("abstestdir" %in% names(preliminary_options)) {
  setwd(preliminary_options$abstestdir)
}

# Check that any provided GAMS system directory exists
sysdir <- NA
if ("sysdir" %in% names(preliminary_options)) {
  sysdir <- preliminary_options$sysdir
  # Strip any trailing / or \ since file.exists may not like it
  if (str_sub(sysdir, -1, -1) %in% c("/", "\\")) {
    sysdir <- str_sub(sysdir, 1, -2)
  }
  if (!file.exists(sysdir)) {
    stop(str_glue("Invalid sysdir= option value, {sysdir} does not exist!"))
  }
}
   
# Make sure the Excel file exists, unless it is a dummy test argument
if (excel_file != "dummy.xls" && excel_file != "dummy.xlsx") {
  if (!(file.exists(excel_file))) {stop(str_glue("Excel file does not exist!: '{excel_file}'"))}
}


# Make sure that any specified options file exists.
if (!is.na(options_file)) {
  if (!(file.exists(options_file))) {stop(str_glue("Options file does not exist!: '@{options_file}'"))}
}

# Use given GDX output file, or set default
if ("output" %in% names(preliminary_options)) {
  gdx_file <- preliminary_options$output
} else {
  gdx_file <- str_c(extensionless_excel_name, ".gdx")
}
rm(extensionless_excel_name)

# ---- Expand options from index sheet or options file ----

more_opts <- c()

# Read options from any index sheet
if ("index" %in% names(preliminary_options)) {
  # Read the index sheet
  cl <- range2cell_limits(preliminary_options$index)
  tib <- suppressMessages(read_excel(excel_file, range=cl))
  # Make sure the column headers are lower case
  colnames(tib) <- str_to_lower(colnames(tib))
  # Require five columns
  col_count <- length(colnames(tib))
  if (col_count != 5) {
    stop(str_glue("Index sheet has {col_count} columns, expecting 5!"))
  }
  # Require  column names
  if (!(all(colnames(tib) == c("...1", "...2", "...3", "rdim", "cdim")) || all(colnames(tib) == c("...1", "...2", "...3", "cdim", "rdim")))) {
    stop(str_glue("Unexpected column names in index sheet. The only supported naming has the first three out of 5 columns unnamed, and the last two columns should be named 'rdim' and 'cdim'."))
  }
  # Turn the tibble rows into strings with key=value options and extract these as a character vector
  tib <- tib %>% transmute(rows=str_glue("{.[[1]]}={.[[2]]} rng={.[[3]]} cdim={cdim} rdim={rdim}"))
  rows <- as.character(tib$rows)
  # Extract the options as strings and append them
  index_opts <- as.character(str_split(str_c(rows, collapse=" "), "[:blank:]+", simplify=TRUE))
  more_opts <- c(more_opts, index_opts)
  # Cleanup
  rm(cl, tib, col_count, rows, index_opts)
}

# Read any options file
if (!is.na(options_file)) {
  of_conn <- file(options_file, open="rt")
  lines <- readLines(of_conn)
  close(of_conn)
  of_opts <- as.character(str_split(str_replace_all(trimws(str_c(lines, collapse=" ")), "[:blank:]*=[:blank:]*", "="), "[:blank:]+", simplify=TRUE))
  more_opts <- c(more_opts, of_opts)
  rm(of_conn, of_opts)
}

rm(preliminary_options)

# ---- Parse and check expanded arguments  ----

# Match possibly expanded (keyword=value) options and get their names and values
option_matches <- str_match(c(args, more_opts), "^([:alpha:]+)=(.*)$")
onames <- str_to_lower(option_matches[,2][!is.na(option_matches[,1])])
values <- option_matches[,3][!is.na(option_matches[,1])]

# Define options classes
PUBLIC_GLOBAL_OPTIONS <- c("index", "output", "sysdir")
GLOBAL_OPTIONS <- c(PUBLIC_GLOBAL_OPTIONS, "testdir", "abstestdir")
SYMBOL_OPTIONS <- c("dset", "par", "set")
SYMBOL_ATTRIBUTE_OPTIONS <- c("cdim", "rdim", "rng", "project")
ALL_OPTIONS <- c(GLOBAL_OPTIONS, SYMBOL_OPTIONS, SYMBOL_ATTRIBUTE_OPTIONS)

# Check that all option names are valid
if (!all(onames %in% ALL_OPTIONS)) {
  stop(str_glue("Invalid option: '{args[[i]]}'!"))
}

# Classify option names and assert that the classes do not intersect and cover all keywords
is_global           <- onames %in% GLOBAL_OPTIONS
is_symbol           <- onames %in% SYMBOL_OPTIONS
is_symbol_attribute <- onames %in% SYMBOL_ATTRIBUTE_OPTIONS
stopifnot(!any(is_global & is_symbol))
stopifnot(!any(is_global & is_symbol_attribute))
stopifnot(!any(is_symbol & is_symbol_attribute))
stopifnot(all(is_global | is_symbol | is_symbol_attribute))

# Check that any global options precede symbol and symbol attribute options
last_global_index <- 0
if (length(onames) > 0) {
  is_global_rl <- rle(is_global)
  if ((any(is_global) && !is_global_rl$values[[1]]) || length(is_global_rl$values) > 2) {
    stop(str_glue("Invalid order of options! Global options must precede any symbol or symbol attribute options."))
  }
  if (any(is_global)) {
    last_global_index <- is_global_rl$lengths[[1]]
  }
}

# Stick the global option names and values into a dictionary
if (length(onames) > 0) {
  global_options <- as.list(structure(values[1:last_global_index], names=onames[1:last_global_index]))
} else {
  global_options <- list()
}

# Check symbol options and their attributes and store them as per-symbol dictionaries
symbol_dicts <- list()
symbol_dict <- NULL
i <- last_global_index+1
while (i <= length(onames)) {
  if (is_symbol[[i]]) {
    # Handle any dictionary from prior symbol
    if (!is.null(symbol_dict)) {
      # Check that dictionary of prior symbol has attributes
      if (!has_attributes) {
        stop(str_glue("Symbol option {symbol_dict$type}={symbol_dict$name} has no symbol attributes!"))
      }
      # Store the dictionary of the prior symbol
      symbol_dicts[[length(symbol_dicts) + 1]] <- symbol_dict
    }
    # Start a new symbol dictionary
    symbol_dict <- list(name=values[[i]], type=onames[[i]])
    has_attributes <- FALSE
  } else {
    # Check that there is a prior symbol to which the attribute belongs
    if (is.null(symbol_dict)) {
      stop(str_glue("Invalid position of option {onames[[i]]}={values[[i]]}! Symbol attribute options must follow a symbol option."))
    }
    # Check that the attribute is the first of its kind for the symbol
    if (onames[[i]] %in% names(symbol_dict)) {
      stop(str_glue("Multiple {onames[[i]]} attributes for symbol option {symbol_dict$type}={symbol_dict$name}!"))
    }
    # Stick the attribute into the symbol dictionary
    symbol_dict[[onames[i]]] <- values[[i]]
    has_attributes <- TRUE
  }
  i <- i + 1
}
# Handle dictionary of last symbol, if any
if (!is.null(symbol_dict)) {
  # Check that symbol dictionary has attributes
  if (!has_attributes) {
    stop(str_glue("Symbol option {symbol_dict$type}={symbol_dict$name} has no symbol attributes!"))
  }
  # Store the dictionary of the last symbol
  symbol_dicts[[length(symbol_dicts) + 1]] <- symbol_dict
}

# Fail when there are no symbol options with which to perform the conversion
if (length(symbol_dicts) == 0) {
  stop("No symbol options, cannot perform conversion!")
}

# Check rng symbol attributes and convert to cell_limits objects
for (i in 1:length(symbol_dicts)) {
  # Retrieve symbol dictionary from list
  symbol_dict <- symbol_dicts[[i]]
  if ("rng" %in% names(symbol_dict)) {
    # Convert the range string to cell_limits
    cl <- range2cell_limits(symbol_dict$rng)
    # Update symbol dictionary and return it to list
    suppressWarnings(symbol_dict$rng <- cl)
    symbol_dicts[[i]] <- symbol_dict
  }
}

# Coerce cdim and rdim symbol attributes to integers
for (i in 1:length(symbol_dicts)) {
  # Retrieve symbol dictionary from list
  symbol_dict <- symbol_dicts[[i]]
  # Coerce any cdim field to integer
  if ("cdim" %in% names(symbol_dict)) {
    cdim <- symbol_dict$cdim
    suppressWarnings(cdim <- as.integer(cdim))
    if (is.na(cdim)) {stop(str_glue("Non-integer cdim option value for symbol {symbol_dict$name}!"))}
    if (cdim < 1) {stop(str_glue("Invalid cdim={cdim} option value for symbol {symbol_dict$name}!"))}
    symbol_dict$cdim <- cdim
  }
  # Coerce any rdim field to integer
  if ("rdim" %in% names(symbol_dict)) {
    rdim <- symbol_dict$rdim
    suppressWarnings(rdim <- as.integer(rdim))
    if (is.na(rdim)) {stop(str_glue("Non-integer rdim option value for symbol {symbol_dict$name}!"))}
    if (rdim < 1) {stop(str_glue("Invalid rdim={rdim} option value for symbol {symbol_dict$name}!"))}
    symbol_dict$rdim <- rdim
  }
  # Return updated symbol dictionary to list
  symbol_dicts[[i]] <- symbol_dict
}

# Check project symbol attributes
for (i in 1:length(symbol_dicts)) {
  # Retrieve symbol dictionary from list
  symbol_dict <- symbol_dicts[[i]]
  if ("project" %in% names(symbol_dict)) {
    if (symbol_dict$type != "par") {
      stop(str_glue("Project option not supported for symbol {symbol_dict$type}={symbol_dict$name}: only supported for par symbols!"))
    }
    if (!(symbol_dict$project %in% c('Y', 'N'))) {
      stop(str_glue("Invalid project option value '{symbol_dict$project}' for symbol {symbol_dict$name}!"))
    }
  }
}

# Clean up from argument parsing
rm(is_global, is_symbol, is_symbol_attribute, last_global_index, option_matches, onames, symbol_dict, values)

# ---- Make sure the GDX libraries are loaded ----

if (is.na(sysdir)) {
  # Use GAMS system directory passed via PATH/[DY]LD_LIBARRY_PATH/R_GAMS_SYSDIR to load the GDX libraries for gdxrrw
  if (!igdx(gamsSysDir="", silent=TRUE)) {
    stop("Cannot load GDX libraries! Use the sysdir option or set the GAMS system directory in your PATH (Windows), LD_LIBRARY_PATH (Linux), DYLD_LIBRARY_PATH (macOS) or R_GAMS_SYSDIR (all platforms) environment variable.")
  }
} else {
  # Use GAMS system directory passed via sysdir to load the GDX libraries for gdxrrw
  if (!igdx(gamsSysDir=sysdir, silent=TRUE)) {
    stop(str_glue("Cannot load GDX libraries from provided sysdir {sysdir}"))
  }
}
rm(global_options)

# ---- Convert and write symbols ----

cat(str_glue("xl2gdx {VERSION}, {DATE}"), sep='\n')
cat(str_glue("Input file : {suppressWarnings(normalizePath(excel_file))}"), sep='\n')
cat(str_glue("Output file : {suppressWarnings(normalizePath(gdx_file))}"), sep='\n')

out_list <- list()
for (symbol_dict in symbol_dicts) {
  name    <- symbol_dict$name
  type    <- symbol_dict$type
  rng     <- symbol_dict$rng
  cdim    <- symbol_dict$cdim
  rdim    <- symbol_dict$rdim
  project <- symbol_dict$project

  # ---- par: convert Excel content to GDX parameter via wgdx.reshape ----
  
  if (type == "par") {
  
    if (is.null(cdim)) {stop(str_glue("Missing cdim attribute for symbol {type}={name}"))}  
    if (is.null(rdim)) {stop(str_glue("Missing rdim attribute for symbol {type}={name}"))}  

    # Read Excel subset as a tibble, merging multiple column header rows if needed
    # NOTE: yields UTF-8 strings in case of special characters
    if (cdim == 1) {
      tib <- suppressMessages(read_excel(excel_file, range=rng, guess_max=GUESS_MAX, trim_ws=TRIM_WS))
      col_names <- colnames(tib)
      # Cut-off any columns as of first empty in-range column like GDXXRW does
      for (col in 1:length(tib)) {
        if (col_names[[col]] == str_c("...", col)) {
          # Column has no name
          if (all(is.na(tib[[col]]))) {
            # Column has no values either, cut it and all columns to the right off
            tib <- select(tib, -(col:length(tib)))
            col_names <- colnames(tib)
            break
          }
        }
      }
    } else {
      stopifnot(cdim > 1)
      # Multiple column header rows, read them first
      header_row_rng <- rng
      header_row_rng$lr[[1]] <- rng$ul[[1]]+cdim-1
      col_header_rows <- suppressMessages(read_excel(excel_file, col_names=FALSE, range=header_row_rng))
      # Merge the rows with a <#> separator to column names
      col_names <- apply(col_header_rows, 2, function(col) str_c(col, collapse="<#>"))
      # Read the range without the header rows, instead setting the merged colulumn names
      headerless_rng <- rng
      headerless_rng$ul[[1]] <- rng$ul[[1]]+cdim
      if (is.na(rng$lr[[2]])) {
        # Open-ended range of columns, make sure to read as many as extracted column names
        headerless_rng$lr[[2]] <- rng$ul[[2]] + length(col_names) -1
      }
      tib <- suppressMessages(read_excel(excel_file, col_names=col_names, range=headerless_rng, guess_max=GUESS_MAX, trim_ws=TRIM_WS))
      rm(header_row_rng, col_header_rows, headerless_rng)
    }

    # Check that sufficient columns were read to satisfy rdim
    if (length(tib) < rdim+1) {
      stop(str_glue("Too few columns in Excel input of symbol {type}={name}, should be at least rdim+1 since there must be at least one value column!"))
    }

    # Check whether column names are valid
    if (typeof(col_names) != "character") {
      stop("Extracted column names are not character strings!")
    }
    if (any(Encoding(col_names) == "UTF-8")) {
      stop(str_c("Special characters in column names not supported!: ", str_c(col_names[Encoding(col_names) == "UTF-8"], collapse=", "), collapse=""))
    }
    col_named <- col_names != str_c("...", as.character(1:length(tib)))
    if (length(tib) > rdim+1) {
      # Multiple value columns
      if (!all(col_named[(rdim+1):length(tib)])) {
        stop(str_glue("Excel input of symbol {type}={name} has multiple value columns ({rdim+1}-{length(tib)}), but not all these columns have header names and as such cannot be gathered to a gdx dimension!"))
      }
      if (length(unique(col_names[(rdim+1):length(tib)])) != length(tib)-rdim) {
        stop(str_glue("Excel input of symbol {type}={name} has multiple value columns ({rdim+1}-{length(tib)}), but their header names are not unique as such cannot be gathered to a gdx dimension!"))
      }
    }
    
    # Project latin special characters in non-value columns to ASCII.
    # Unlike iconv(), stri_trans_general() yields the same results independent of locale and OS.
    to_latin <- FALSE
    for (col in 1:rdim) {
      if (typeof(tib[[col]]) == "character") {
        # A character column, for efficiency first collect the unique strings
        uniq <- unique(tib[[col]])
        if (any(Encoding(uniq) == "UTF-8")) {
          if (!is.null(project) && project == 'Y') {
            # Check that unique column strings can be projected to ASCII
            uniq_proj <- stri_trans_general(uniq, "Latin-ASCII")
            if (any(Encoding(uniq_proj) == "UTF-8")) {
              stop(str_c("Cannot project special characters to ASCII: ", str_c(uniq_proj[Encoding(uniq_proj) == "UTF-8"], collapse=", "), collapse=""))
            }
            # Project column to ASCII
            message(str_c("Special characters projected to ASCII look-alikes: ", str_c(str_c(uniq[Encoding(uniq) == "UTF-8"], uniq_proj[Encoding(uniq) == "UTF-8"], sep=" -> "), collapse=", "), collapse=""))
            tib[[col]] <- stri_trans_general(tib[[col]], "Latin-ASCII")
            rm(uniq_proj)
          } else {
            # Check that unique column strings can be respresented as Windows-1252 (latin code page)
            uniq_rep <- stri_encode(uniq, from="UTF-8", to="Windows-1252")
            if (any(Encoding(uniq_rep) == "UTF-8")) {
              stop(str_c("Cannot represent special characters with Windows-1252 (ASCII extended with latin code page): ", str_c(uniq_rep[Encoding(uniq_rep) == "UTF-8"], collapse=", "), collapse=""))
            }
            # Represent column as Windows-1252 (ASCII extended with latin code page)
            to_latin <- TRUE
            tib[[col]] <- stri_encode(tib[[col]], from="UTF-8", to="Windows-1252")
            rm(uniq_rep)
          }
        }
        rm(uniq)
      }
    }
    if (to_latin) {
      cat(str_glue("Note: non-ASCII special characters are present for symbol {type}={name} in column {col}. These were represented as Windows-1252 (ASCII extended with a latin code page). Handling of such text is locale-sensitive. Consider to project this symbol to ASCII using project=Y so that you can use locale-insensitive pure-ASCII handling after loading the GDX."), sep='\n')
    }

    # Prepare tibble
    if (length(tib) == rdim+1 && !col_named[[rdim+1]]) {
      # A single unnamed value column, no gathering required, only drop rows with NA's
      tib <- tib %>% drop_na()
    } else {
      # Gather value column or columns using wgdx.reshape or dplyr
      if (RESHAPE) {
        # Reshape to collect value columns and add to list of symbols to output, does its own factoring
        tib <- wgdx.reshape(tib, rdim+1, symName=name, setsToo=FALSE)[[1]] %>% drop_na
      } else {
        # Gather value-containing columns as a new pair of key-value columns
        tib <- tib %>%
          gather(col_names[(rdim+1):length(col_names)], key="gathered_keys", value="gathered_values", na.rm=TRUE)
        
        # Factor the keys gathered from the value column headers
        tib$gathered_keys <- factor(tib$gathered_keys)
      }
    }

    # Factor non-value columns where needed
    for (col in 1:rdim) {
      if (!is.factor(tib[[col]])) {tib[[col]] <- factor(tib[[col]])}
    }
    
    # Work around Excel having converted number-alike strings to binary floating point representation, thereby introducing
    # rounding errors of fractional decimal values that cannot be represented exactly as a binary floating point number.
    for (col in 1:rdim) {
      lvls <- levels(tib[[col]])
      if (typeof(lvls) == "character") {
        ma <- str_match(lvls, "[.][:digit:]+[09]{8}[:digit:]")
        if (!all(is.na(ma))) {
          # Occurrences of >=8 consecutive 0s or 9s after a point, Excel mangling is likely, try to convert these to double
          dbls <- suppressWarnings(as.double(lvls[!is.na(ma)]))
          if (!all(is.na(dbls))) {
            warning(str_glue("Fixing Excel mangling for symbol {type}={name} column {col}."))
            # Revert converted doubles to character strings, getting rid of binary rounding through decimal rounding
            lvls[!is.na(ma)][!is.na(dbls)] <- as.character(dbls[!is.na(dbls)])
            # Replace with fixed levels
            levels(tib[[col]]) <- lvls
          }
          rm(dbls)
        }
        rm(ma)
      }
      rm(lvls)
    }

    if (cdim > 1) {
      # Separate gathered column into separate columns, one for each column header row
      tib <- separate(tib, rdim+1, into=str_c("...", (rdim+1):(rdim+cdim)), sep="<#>")
      # Factor separated columns where needed
      for (col in (rdim+1):(rdim+cdim)) {
        if (!is.factor(tib[[col]])) {tib[[col]] <- factor(tib[[col]])}
      }
    }
    
    # Annotate and add tibble to output list
    attr(tib, "symName") <- name
    attr(tib, "ts") <- str_glue("Converted from {basename(excel_file)}{ifelse(is.na(rng$sheet), '', str_glue(' sheet {rng$sheet}'))}")
    out_list[[length(out_list)+1]] <- tib
    rm(tib)
  }

  # ---- dset: convert Excel content to GDX set ----
  
  if (type == "dset") {
  
    if (!is.null(cdim)) {stop("A cdim option is not yet supported when using the dset option!")}  
    if (is.null(rdim)) {stop(str_glue("Missing rdim attribute for symbol {type}={name}"))}  
    if (rdim != 1) {stop("Only cdim=1 is allowed when using the dset option!")}
    rng$lr <- c(NA, rng$ul[[2]])
  
    # Read Excel subset as a tibble
    # NOTE: yields UTF-8 strings in case of special characters
    # NOTE: trims leading and trailing whitespace
    tib <- suppressMessages(read_excel(excel_file, range=rng, col_names=FALSE))
  
    t <- tib[[1]] %>% sort %>% unique
  
    # Add to output list
    l <- list(name=name,
              type="set",
              dim=1,
              form="full",
              ts=str_glue("Converted from {basename(excel_file)}{ifelse(is.na(rng$sheet), '', str_glue(' sheet {rng$sheet}'))}"),
              uels=c(list(c(t)))
              )
    out_list[[length(out_list)+1]] <- l
    rm(tib, l)
  }

}

# Write the symbols
wgdx.lst(gdx_file, out_list)

# Print total processing time
cat(str_glue("Total time = {format(Sys.time()-start_time)}"), sep='\n')
