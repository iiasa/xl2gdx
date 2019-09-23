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
# gdxrrw R package: https://www.gams.com/latest/docs/T_GDXRRW.html
# tidyverse R package collection: https://www.tidyverse.org/
#
# BEWARE, on Windows installing the gdxrrw source package will not work unless
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
# BEWARE, leading and trailing whitespace in fields is trimmed and special
# characters in the Excel file are projected onto their closest ASCII
# look-alikes so as to avoid later character encoding issues. By limiting
# to ASCII, locale and platform dependencies are avoided. This is UNLIKE
# GDXRRW which stores special characters to the GDX in what appears to be a
# locale-dependent encoding.
#
# Author: Albert Brouwer
#
# Todo:
# - test A1 conversions w/o sheet spec and w single sheet in XL
# - support index=, get options from an Excel sheet (one row per symbol, presumably)
# - Both reshape TRUE/FALSE write 100000/200000 as 1e+05 2e+05
# - support the clear symbol attribute
# - support set=
# - support skipempty=0, at least in conjunction with index=
# - combine RESHAPE TRUE/FALSE where possible

options(tidyverse.quiet=TRUE)
library(gdxrrw)
library(tidyverse)
library(cellranger) # installed when you install tidyverse
library(readxl) # installed when you install tidyverse
library(stringi) # installed when you install tidyverse
RESHAPE <- TRUE # select wgdx.reshape (TRUE) or dplyr-based (FALSE) parameter writing

# ---- Get command line arguments, or provide test arguments when running from RStudio ----

if (Sys.getenv("RSTUDIO") == "1") {
  # Failing argument parsing test cases
  #args <- c() # no arguments, error with usage
  #args <- c("output=foo") # no Excel file as first argument
  #args <- c("@options_file") # no Excel file as first argument
  #args <- c("does_not_exist.xls") # not-existent Excel file
  #args <- c("dummy.xls") # an xls, but no symbol.
  #args <- c("dummy.xlsx") # an xlsx, but no symbol
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
  args <- c("dummy.xlsx", "par=foo", "rng=A1") # missing cdim attribute for par
  #args <- c("dummy.xlsx", "par=foo", "rng=A1", "cdim=1") # missing rdim attribute for par
  #args <- c("dummy.xlsx", "dset=foo", "rng=A1") # missing rdim attribute for dset
  #args <- c("dummy.xlsx", "set=foo", "rng=bar!A1:B2") # no end col/row allowed for a set
  #args <- c("dummy.xlsx", "par=foo", "rng=bar!A1:B2", "cdim=invalid") # non-integer cdim
  #args <- c("dummy.xlsx", "par=foo", "rng=bar!A1:B2", "rdim=invalid") # non-integer rdim
  
  # Conversion tests
  #args <- c("test.xls",  "testdir=test1", "par=para",   "rng=toUse!c4:f39",               "cdim=1", "rdim=1")
  #args <- c("test.xlsx", "testdir=test2", "par=para",   "rng=CommodityBalancesCrops1!a1", "cdim=1", "rdim=7")
  #args <- c("test.xlsx", "testdir=test3", "dset=doset", "rng=TradeSTAT_LiveAnimals1!f2",            "rdim=1")
  #args <- c("test.xlsx", "testdir=test4", "par=para",   "rng=Sheet1!AV2:BA226",           "cdim=1", "rdim=2", "par=parb", "rng=Sheet1!B2:AT226", "cdim=1", "rdim=2")
  #args <- c("test.xlsx", "testdir=test5", "par=para",   "rng=A1",                         "cdim=1", "rdim=1")
  #args <- c("test.xls",  "testdir=test6", "@taskin1.txt")
} else {
  args <- commandArgs(trailingOnly=TRUE)
}

# ---- Display usage if needed ----

USAGE <- str_c("Usage:",
              "Rscript xl2gdx.R <Excel file> [options] [@<options file>] [symbols]",
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

# Ensure that the first argument is an Excel file
if (str_sub(args[[1]], 1, 1) == "@" || !is.na(option_matches[[1, 1]])) {
  stop("First argument must be an Excel file!")
}
excel_file <- args[[1]]
excel_base_path <- str_match(excel_file, "^(.+)[.][xX][lL][sS][xX]?$")[2]
if (is.na(excel_base_path)) {stop(str_glue("Not an Excel file: absent .xls or .xlsx extension in first argument '{excel_file}'!"))}

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

# Change current directory when testing from RStudio
if ("testdir" %in% names(preliminary_options)) {
  setwd(str_c(dirname(rstudioapi::getActiveDocumentContext()$path), "/", preliminary_options$testdir))
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
  gdx_file <- str_c(excel_base_path, ".gdx")
}
rm(excel_base_path)

# ---- Expand options from index sheet or options file ----

more_opts <- c()

#TODO: index= handling goes here
if ("index" %in% names(preliminary_options)) {
  cl <- range2cell_limits(preliminary_options$index)
  tib <- suppressMessages(read_excel("test.xls", range=cl))
  rm (cl, tib)
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

# ---- Parse and check expanded arguments  ----

# Match possibly expanded (keyword=value) options and get their names and values
option_matches <- str_match(c(args, more_opts), "^([:alpha:]+)=(.*)$")
onames <- str_to_lower(option_matches[,2][!is.na(option_matches[,1])])
values <- option_matches[,3][!is.na(option_matches[,1])]

# Define options classes
PUBLIC_GLOBAL_OPTIONS <- c("index", "output", "sysdir")
GLOBAL_OPTIONS <- c(PUBLIC_GLOBAL_OPTIONS, "testdir")
SYMBOL_OPTIONS <- c("dset", "par", "set")
SYMBOL_ATTRIBUTE_OPTIONS <- c("cdim", "rdim", "rng")
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
  if ("rng" %in% names(symbol_dict)) {
    # Retrieve symbol dictionary from list
    symbol_dict <- symbol_dicts[[i]]
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
    if (is.na(cdim)) {stop("Non-integer cdim option value for symbol {symbol_dict$name}!")}
    symbol_dict$cdim <- cdim
  }
  # Coerce any rdim field to integer
  if ("rdim" %in% names(symbol_dict)) {
    rdim <- symbol_dict$rdim
    suppressWarnings(rdim <- as.integer(rdim))
    if (is.na(rdim)) {stop("Non-integer rdim option value for symbol {symbol_dict$name}!")}
    symbol_dict$rdim <- rdim
  }
  # Return updated symbol dictionary to list
  symbol_dicts[[i]] <- symbol_dict
}

# Clean up from argument parsing
rm(is_global, is_symbol, is_symbol_attribute, option_matches, onames, values)

# ---- Make sure the GDX libraries are loaded ----

if ("sysdir" %in% names(global_options)) {
  # Use GAMS system directory passed via sysdir to load the GDX libraries for gdxrrw
  if (!igdx(gamsSysDir=global_options$sysdir, silent=TRUE)) {
    stop(str_glue("Cannot load GDX libraries from provided sysdir {global_options$sysdir}"))
  }
} else {
  # Use GAMS system directory passed via PATH/[DY]LD_LIBARRY_PATH/R_GAMS_SYSDIR to load the GDX libraries for gdxrrw
  if (!igdx(gamsSysDir="", silent=TRUE)) {
    stop("Cannot load GDX libraries! Use the sysdir option or set the GAMS system directory in your PATH (Windows), LD_LIBRARY_PATH (Linux), DYLD_LIBRARY_PATH (macOS) or R_GAMS_SYSDIR (all platforms) environment variable.")
  }
}

# ---- Convert and write symbols ----

out_list <- list()
for (symbol_dict in symbol_dicts) {
  name <- symbol_dict$name
  type <- symbol_dict$type
  rng  <- symbol_dict$rng
  cdim <- symbol_dict$cdim
  rdim <- symbol_dict$rdim

  # ---- par: convert Excel content to GDX parameter via wgdx.reshape ----
  
  if (type == "par" && RESHAPE) {
  
    if (is.null(cdim)) {stop(str_glue("Missing cdim attribute for symbol {type}={name}"))}  
    if (is.null(rdim)) {stop(str_glue("Missing rdim attribute for symbol {type}={name}"))}  
    if (cdim != 1) {stop("cdim != 1 not yet supported when using the par option!")}
    
    # Read Excel subset as a tibble
    # NOTE: yields UTF-8 strings in case of special characters
    # NOTE: trims leading and trailing whitespace
    tib <- suppressMessages(read_excel(excel_file, range=rng))
    
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
            stop(str_c("Cannot project special characters to ASCII: ", str_c(unipro[Encoding(uniq_proj) == "UTF-8"], collapse=", "), collapse=""))
          }
          warning(str_c("Special characters projected to ASCII look-alikes: ", str_c(str_c(uniq[Encoding(uniq) == "UTF-8"], uniq_proj[Encoding(uniq) == "UTF-8"], sep=" -> "), collapse=", "), collapse=""))
          proj <- stri_trans_general(tib[[r]], "latin-ascii")
          tib[[r]] <- proj
        }
      }
    }
  
    # Reshape to collect value columns and add to list of symbols to output
    attr(tib, "ts") <- str_glue("Converted from {basename(excel_file)}{ifelse(is.na(rng$sheet), '', str_glue(' sheet {rng$sheet}'))}")
    lst <- wgdx.reshape(tib, rdim+1, symName=name, setsToo=FALSE)[[1]] %>% drop_na
    out_list[[length(out_list)+1]] <- lst
  }
  
  # ---- par: convert Excel content to GDX parameter via tibble fu ----
  
  if (type == "par" && !RESHAPE) {
  
    if (is.null(cdim)) {stop("Missing cdim option, must be present when using the par option!")}  
    if (is.null(rdim)) {stop("Missing rdim option, must be present when using the par option!")}  
    if (cdim != 1) {stop("cdim != 1 not yet supported when using the par option!")}
  
    # Read Excel subset as a tibble
    # NOTE: yields UTF-8 strings in case of special characters
    # NOTE: trims leading and trailing whitespace
    tib <- suppressMessages(read_excel(excel_file, range=rng))
    
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
            stop(str_c("Cannot project special characters to ASCII: ", str_c(unipro[Encoding(uniq_proj) == "UTF-8"], collapse=", "), collapse=""))
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
  
    # Gather value-containing columns as a new pair of key-value columns
    g <- tib %>%
         gather(col_names[(rdim+1):length(col_names)], key="gathered_keys", value="gathered_values", na.rm=TRUE)
  
    # Factor the keys gathered from the value column headers
    g$gathered_keys <- factor(g$gathered_keys)
  
    # Add to output list
    #attr(g, "domains") <- col_names[1:rdim] # This sets the column names as domains
    attr(g, "symName") <- name
    attr(g, "ts") <- str_glue("Converted from {basename(excel_file)}{ifelse(is.na(rng$sheet), '', str_glue(' sheet {rng$sheet}'))}")
    out_list[[length(out_list)+1]] <- g
  }
  
  # ---- dset: convert Excel content to GDX set ----
  
  if (type == "dset") {
  
    if (!is.null(cdim)) {stop("A cdim option is not yet supported when using the dset option!")}  
    if (is.null(cdim)) {stop(str_glue("Missing cdim attribute for symbol {type}={name}"))}  
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
  }

}
wgdx.lst(gdx_file, out_list)
