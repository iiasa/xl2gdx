# xl2gdx

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

## Requirements:

- [**gdxrrw**](https://github.com/GAMS-dev/gdxrrw), an R package for
  reading/writing GDX files from R. For a list of which package versions
  match what R versions, see the [**gdxrrw** wiki](https://github.com/GAMS-dev/gdxrrw/wiki).
- The [tidyverse](https://www.tidyverse.org/) curated R package collection.

**Beware** as of version V1.0.8, **gdxrrw** requires GAMS >= V33.
When you use an earlier GAMS version, use an earlier **gdxrrw** version.
