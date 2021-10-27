# xl2gdx.R

R script to convert Excel to [GDX](https://www.gams.com/latest/docs/UG_GDX.html):
- Can replace GDXXRW for Excel-to-GDX conversion.
- Accepts the same arguments and a subset of the options that GDXXRW does.
- Unlike GDXXRW, works on non-Windows platforms and does not require Office.

Tests are located in the separate private [xl2gdx-tests](https://github.com/iiasa/xl2gdx-tests)
repository. That repository is private because the the licensing conditions of
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
- [R](https://www.r-project.org). After installation, ensure that `Rscript` is
  on-path.
- The [tidyverse](https://www.tidyverse.org/) curated R package collection.
- [**gdxrrw**](https://github.com/GAMS-dev/gdxrrw), an R package for
  reading/writing GDX files from R. To
  [make **gdxrrw** find the GAMS system directory](https://github.com/GAMS-dev/gdxrrw#checking-if-gdxrrw-is-installed-correctly)
  you can use the `sysdir` command line option (see below) or make sure a
  sufficiently recent GAMS installation directory is included in either the
  `PATH` (on Windows) or `LD_LIBRARY_PATH` (on Linux) environment variable.
  To make things more explicit, you can instead point the **gdxrrw**-specific
  environment variable `R_GAMS_SYSDIR` to a GAMS installation directory. It is
  probably best to choose the most recent version of GAMS that is installed.

  * If you use an environment variable to point to the GAMS installation
    directory, the following should work and report the used environment
    variable:
    ```
    $ R
    > library(gdxrrw)
    > igdx(gamsSysDir='')
    ```
  * **Beware**, recent versions of **gdxrrw** use a new GDX
    [API](https://en.wikipedia.org/wiki/API)
    that is only available with recent versions of GAMS. Make sure to
    [check the installation](https://github.com/GAMS-dev/gdxrrw#checking-if-gdxrrw-is-installed-correctly).
    If need be, install a newer GAMS version.
  * On Windows, it will likely prevent problems when you first install
    [Rtools](https://cran.r-project.org/bin/windows/Rtools/)
    so that you can compile the **gdxrrw** and other R packages from source.
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
- `rng='[<sheet>!]<start_colrow>[:<end_colrow>]'`
- `project=Y` (project latin special characters to ASCII for `par=` symbols,
  defaults to `N`)

### project_to_ASCII.R
  
To invoke `project_to_ASCII.R`, issue:

`[Rscript ]project_to_ASCII.R <text file with special characters>`

This projects the given text file to ASCII when possible, replacing it
in-place.
