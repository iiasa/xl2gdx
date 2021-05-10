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

None. This is not an R package, `xl2gdx.R` is a utility script that can be invoked with
command line parameters. Just copy it to a handy location. The same holds for the
`project_to_ASCII.R` helper script.

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

`project_to_ASCII.R ` depends on:
- The [tidyverse](https://www.tidyverse.org/) curated R package collection.
