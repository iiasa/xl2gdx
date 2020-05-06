# xl2gdx

R script to convert Excel to GDX:
- Can replace GDXXRW for Excel-to-GDX conversion.
- Accepts the same arguments and a subset of the options that GDXXRW does.
- Unlike GDXXRW, works on non-Windows platforms and does not require Office.

Note that the included tests are Windows-only and do require Office since
they compare GDXXRW output to xl2gdx output. The licensing conditions of
the input Excel sheets included with the tests were not reviewed. Hence
this repository is private: do not redistribute the test data. It is fine
to redistribute the `xl2gdx.R` script.

For further information read the header comments in the script and see
the [GDXXRW documentation](https://www.gams.com/latest/docs/T_GDXXRW.html).

## Requirements:

- [gdxrrw](https://www.gams.com/latest/docs/T_GDXRRW.html), an R package
  for reading/writing GDX files from R.
- The [tidyverse](https://www.tidyverse.org/) curated R package collection.
