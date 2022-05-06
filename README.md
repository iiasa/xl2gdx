# xl2gdx.R

R script to convert Excel to [GDX](https://www.gams.com/latest/docs/UG_GDX.html):
- Can replace [GDXXRW](https://www.gams.com/latest/docs/T_GDXXRW.html) for
  Excel-to-GDX conversion.
- Accepts the same arguments and a subset of the options that GDXXRW does.
- Unlike GDXXRW, works on non-Windows platforms and does not require Office.

Tests are located in the separate private [xl2gdx-tests](https://github.com/iiasa/xl2gdx-tests)
repository. That repository is private because the licensing conditions of
the corner-case input Excel sheets included with the tests are diverse and
were not reviewed. To request access to the tests repository, email the
author or post an issue in the issue tracker.

For further information read the header comments in the script and see
the [GDXXRW documentation](https://www.gams.com/latest/docs/T_GDXXRW.html).

## Installation

This is not an R package, instead `xl2gdx.R` is a utility script that can be
invoked with command line parameters. Just copy it to a handy location. The
same holds for the `project_to_ASCII.R` helper script. The dependencies listed
below should first be installed though, and some environment variable may need
to be set.

## Dependencies

`xl2gdx.R` depends on:
- [R](https://www.r-project.org). After installation, ensure that `R` and `Rscript` are
  on-path by adding the right installation subdirectory to the `PATH` environment variable.
  On Windows, this directory ends in `R-x.y.z\bin\x64` for the 64-bit binaries, where
  `x.y.z` is the R version.
  
  * **⚠️Warning:** When installing R, old R versions are not automatically removed. Having
  multiple R versions installed can cause confusion. Remove the older version unless you
  have good reasons to keep them.
  
  * **⚠️Note:** After updating R, you will need re-install R packages.

  * **⚠️Attention:** When you use RStudio and update R, you should make sure that RStudio
  is using the new R installation by configuring it under **Tools >> Global options ... >> General**.

- The [tidyverse](https://www.tidyverse.org/) curated R package collection. From the R prompts, install with
  ```R
  install.packages("tidyverse")
  ```
  or use the RStudio package manager.
- [**gdxrrw**](https://github.com/GAMS-dev/gdxrrw), an R package for
  reading/writing GDX files from R. To
  [make **gdxrrw** find the GAMS system directory](https://github.com/GAMS-dev/gdxrrw#checking-if-gdxrrw-is-installed-correctly)
  containing the GDX libraries that it needs to read/write GDX files, you
  can use the `sysdir` command line option (see below) or make sure a
  sufficiently recent GAMS installation directory is included in either the
  `PATH` (on Windows), or `LD_LIBRARY_PATH` (on Linux), or `DYLD_LIBRARY_PATH`
  (on MacOS) environment variable.
  
  It is recommended to instead make things perfectly explicit by setting the
  **gdxrrw**-specific environment variable `R_GAMS_SYSDIR` to point to
  a GAMS installation directory. For reasons explained below, it is best
  to point to the most recent version of GAMS that is installed.
  [See here](https://iiasa.github.io/GLOBIOM/R.html#setting-environment-variables)
  for guidance on how to set environment variables.
  * **⚠️Beware:** changed environment variables are not picked up until you
    restart a process. Therefore, after changing one of the above-mentioned
    environment variables, first restart your command prompt, shell, GAMS
    IDE or GAMS Studio before testing the installation or invoking
    `xl2gdx.R`.
  * If you use an environment variable to point to the GAMS installation
    directory, the following should work and report the used environment
    variable:
    ```R
    $ R
    > library(gdxrrw)
    > igdx(gamsSysDir='')
    ```
  * **⚠️Warning:**, the above will result in an error with recent versions of **gdxrrw** unless you point
    **gdxxrrw** at a GAMS 33 or newer installation directory as per the above instructions. The reason is
    that **gdxrrw** has switched to using an improved GDX [API](https://en.wikipedia.org/wiki/API)
    that is available as of GAMS 33. You may therefore need to install a newer GAMS version
    and point *gdxrrw** at it as per the above instructions.
  * On Windows, it will likely prevent problems when you first
    [install Rtools](https://cran.r-project.org/bin/windows/Rtools/)
    so that you can compile the **gdxrrw** and other R packages from source.
    
    **Beware:** when installing RTools 4.0, do not skip the **Putting Rtools on the PATH** step
    listed in [its installation instructions](https://cran.r-project.org/bin/windows/Rtools/rtools40.html).
  * Without a compiler, you should download a binary **gdxrrw** package
    that matches your R version. For a list of which binary package versions
    match what R versions, see the [**gdxrrw** wiki](https://github.com/GAMS-dev/gdxrrw/wiki).

`project_to_ASCII.R` depends on:
- The [tidyverse](https://www.tidyverse.org/) curated R package collection.

## Usage

Both `xl2gdx.R` and `project_to_ASCII.R` can be invoked via the
[`Rscript`](https://stat.ethz.ch/R-manual/R-devel/library/utils/html/Rscript.html) utility.
It is recommended to add the directory containing `Rscript` to your `PATH`
environment variable so that you can invoke it directly. When you do so on
Linux/MacOS, you can omit the leading `Rscript` from the shell invocation of
the scripts since `Rscript` will then be invoked via the
[shebang header](https://en.wikipedia.org/wiki/Shebang_(Unix))
present in both scripts.

When replacing a GDXXRW invocation in your GAMS code with `xl2gdx.R`, you will typically
have a
[`$call`](https://www.gams.com/36/docs/UG_DollarControlOptions.html#DOLLARcall) or
[`execute`](https://www.gams.com/latest/docs/UG_GamsCall.html#UG_DollarExecute) statement
that invokes GDXXRW. Unless unsupported options are used, it should be possible to
replace the `GDXXRW` or `<path to GAMS dir>/GDXXRW` part of that invocation with
`Rscript <relative path to xl2gdx.R>/xl2gdx.R` and things should work. To verify,
the output of both invocations can be compared with
[`GDXDIFF`](https://www.gams.com/36/docs/T_GDXDIFF.html?search=gdxdiff).

### xl2gdx.R

To invoke `xl2gdx.R` from the command line or shell, issue:

`[Rscript ]xl2gdx.R <Excel file> [options] [@<options file>] [symbols]`

The supported options are listed below. Details for most options are given in
the [GDXXRW manual](https://www.gams.com/latest/docs/T_GDXXRW.html).

#### Global options (provide these first):

- `output=<GDX file>` (if omitted, output to `<Excel file>` but with a `.gdx`
  extension)
- `index='<sheet>!<start_colrow>'`
- `sysdir=<GAMS system directory>`. When omitted, the GAMS installation
  directory must be reachable via an environment variable
  ([see above](#dependencies)).
- `maxdupeerrors=<max>`

#### Symbol options (one or more):

- `dset=<name of domain set to write>`
- `par=<name of parameter to write>`
- `set=<name of set to write>`

#### Symbol attribute options (associated with preceeding symbol):

- `cdim=<number of column dimensions>`
- `rdim=<number of row dimensions>`
- `rng='[<sheet>!]<start_colrow>[:<end_colrow>]'` **⚠️Beware:** unlike GDXXRW
  sheet names are case sensitive.
- `project=Y` (project latin special characters to ASCII for `par=` symbols,
  defaults to `N`)

### project_to_ASCII.R

Project a windows-1252 or ISO-8859-1 encoded text file to ASCII.
Intended to remove special characters from source files. Can be used
in conjunction with he `project=Y` feature of `xl2gdx.R` to locate and
convert special-character references to data in GAMS source files.

**⚠️Warning:** this tool operates in-place, apply it only to source files
under version control so that you can review and revert the changes.

To invoke `project_to_ASCII.R`, issue:

`[Rscript ]project_to_ASCII.R <text file with special characters>`

This projects the given text file to ASCII when possible, replacing it
in-place.

## Troubleshooting

### Error: function 'Rcpp_precious_remove' not provided by package 'Rcpp'

When using `xl2gdx.R` produces this error, upgrade the **Rcpp** package to version 1.0.7 or higher.
