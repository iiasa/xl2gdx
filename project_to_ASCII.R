#!/usr/bin/env Rscript
# Project a windows-1252 or ISO-8859-1 encoded text file to ASCII.
#
# Intended to remove special characters from source files. Can be used
# in conjunction with he project=Y feature of xl2gdx.R to locate and
# convert special-character references to data in GAMS source files.
#
# WARING: this tool operates in-place, apply it only to source files
# under version control so that you can review and revert the changes.
#
# Author: Albert Brouwer

options(tidyverse.quiet=TRUE)
suppressWarnings(library(tidyverse))
suppressWarnings(library(stringi)) # installed when you install tidyverse
suppressWarnings(library(readr)) # installed when you install tidyverse

# ---- Get command line arguments
args <- commandArgs(trailingOnly=TRUE)

# ---- Display usage if needed ----

USAGE <- str_c("Usage:",
              "[Rscript ]project_to_ASCII.R <text file with special characters>",
              "",
              "Projects the given text file to ASCII when possible, replacing it in-place.",
              sep="\n")

# No arguments? Error with usage.
if (length(args) == 0) {
  args <- "trial.orig"
  #stop(str_c("No arguments!", USAGE, sep="\n"))
}

# Display usage if requested
if (args[1] == "?" || args[1] == "-help" || args[1] == "--help") {
  cat(USAGE, sep='\n')
  quit(save="no")
}

# Check that the first argument is an existing file
text_file <- args[[1]]
if (!file.exists(text_file)) {
  stop(str_glue("No such file: '{text_file}'!"))
}

# Set up default locale
loc <- default_locale()

# Read the file as windows-1252 and project to ASCII
loc$encoding <- "windows-1252"
projected <- stri_trans_general(read_file(text_file, locale = loc), "Latin-ASCII")

# If projection has failed, try as ISO-8859-1

if (Encoding(projected) == "UTF-8") {
  loc$encoding <- "ISO-8859-1"
  projected <- stri_trans_general(read_file(text_file, locale = loc), "Latin-ASCII")
  if (Encoding(projected) == "UTF-8") {
    stop("Cannot project special characters to ASCII!")
  }
}

# Save the projected ASCII (subset of UTF-8)
write_file(projected, text_file)
