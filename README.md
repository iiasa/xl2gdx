# xl2gdx.R

R script to convert Excel to GDX:
- Can replace GDXXRW for Excel-to-GDX conversion.
- Accepts the same arguments and a subset of the options that GDXXRW does.
- Unlike GDXXRW, works on non-Windows platforms and does not require Office.

Tests are located in the separate private [xl2gdx-tests](https://github.com/iiasa/xl2gdx-tests)
repository. That repository is private because the the licensing conditions of the
corner-case input Excel sheets included with the tests are diverse and were not reviewed.
To request access to the tests repository, email the author or post an issue
in the issue tracker.

For further information read the header comments in the script and see
the [GDXXRW documentation](https://www.gams.com/latest/docs/T_GDXXRW.html).

## Installation

This is not an R package, instead `xl2gdx.R` is a utility script that can be invoked with
command line parameters. Just copy it to a handy location. The same holds for the
`project_to_ASCII.R` helper script. Noe that the the dependencies listed below should
be installed.

## Dependencies

`xl2gdx.R` depends on:
- The [tidyverse](https://www.tidyverse.org/) curated R package collection.
- [**gdxrrw**](https://github.com/GAMS-dev/gdxrrw), an R package for
  reading/writing GDX files from R. For a list of which binary package versions
  match what R versions, see the [**gdxrrw** wiki](https://github.com/GAMS-dev/gdxrrw/wiki).
  * **Beware**, as of version V1.0.8, **gdxrrw** requires GAMS >= V33.
    When you use an earlier GAMS version, use an earlier **gdxrrw** version.
  * Note that if you can compile packages, for example with [Rtools](https://cran.r-project.org/bin/windows/Rtools/),
    any source package version can be made to work with your R version.
  * If you don't want to go through the hassle of installing Rtools, try a binary
    package built for a slightly earlier R release than the one you have installed.
    A package built for R version x.y.a may work with R version x.y.b (where x, y, a,
    and b are digits and a < b), though possibly with some warnings.

`project_to_ASCII.R` depends on:
- The [tidyverse](https://www.tidyverse.org/) curated R package collection.

## Usage

Both `xl2gdx.R` and `project_to_ASCII.R` can be invoked via the [`Rscript`](https://stat.ethz.ch/R-manual/R-devel/library/utils/html/Rscript.html) utility. It is recommended to add the directory containing `Rscript` to your `PATH` environment variable so that you can invoke it directly, or, on Linux/MacOS, ommit it and have it be located by a [shebang header](https://en.wikipedia.org/wiki/Shebang_(Unix)) present in both scripts.

### xl2gdx.R

To invoke `xl2gdx.R` from the command line or shell, issue:

`[Rscript ]xl2gdx.R <Excel file> [options] [@<options file>] [symbols]`

The supported options are listed below. Details for most options are given in the [GDXXRW manual](https://www.gams.com/latest/docs/T_GDXXRW.html).

#### Global options (provide these first):

- `output=<GDX file>` (if omitted, output to `<Excel file>` but with a `.gdx` extension)
= `index='<sheet>!<start_colrow>'`
= `sysdir=<GAMS system directory>` (pass %gams.sysdir%)
= `maxdupeerrors=<max>`

#### Symbol options (one or more):

- `dset=<name of domain set to write>`
- `par=<name of parameter to write>`
- `set=<name of set to write>`

#### Symbol attribute options (associated with preceeding symbol):

- `cdim=<number of column dimensions>`
- `rdim=<number of row dimensions>`
- `rng='[<sheet>!]<start_colrow>[:<end_colrow>]'`
- `project=Y` (project latin special characters to ASCII for par symbols, defaults to `N`)

### project_to_ASCII.R
  
To invoke `project_to_ASCII.R`, issue:

[Rscript ]project_to_ASCII.R <text file with special characters>

This projects the given text file to ASCII when possible, replacing it in-place.
