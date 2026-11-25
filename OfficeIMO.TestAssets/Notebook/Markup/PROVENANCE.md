# Notebook markup fixtures

These OfficeIMO-authored fixtures reproduce public wire and source contracts.
They contain harmless illustrative code and prewritten output. Validation
reads the supplied content without invoking the document's language runtime.

- `quarto-analysis.qmd` follows the [execution options reference](https://quarto.org/docs/computations/execution-options.html)
  and the [Quarto CLI fixture at 1f794efd](https://github.com/quarto-dev/quarto-cli/blob/1f794efd42ae126344d4100c7fa03829cd23c9cb/tests/docs/smoke-all/2023/01/12/knitr-options-yaml.qmd).
- `myst-analysis.md` follows the [executable Markdown documentation at fe3a7d4d](https://github.com/jupyter-book/mystmd/blob/fe3a7d4de2f1de31af80871944bd63fd9f25583d/docs/notebooks-with-markdown.md).
- `r-analysis.Rmd` follows the [R Markdown notebook source example](https://github.com/rstudio/rmarkdown/blob/9e537b954aded904c34ef755d4c3834132de78e7/tests/testthat/resources/r-notebook.Rmd).
  `r-saved-results.nb.html` embeds this source and follows
  [the annotated HTML Notebook producer at 9e537b95](https://github.com/rstudio/rmarkdown/blob/9e537b954aded904c34ef755d4c3834132de78e7/R/html_notebook.R#L250-L390).
- An independent, archived producer artifact for validation is [Earth Lab's ARINC424.nb.html at 289c9877](https://github.com/earthlab/Navaid-converter/blob/289c9877aaa07b04b650b6e3fa2bfcf1fba974f0/ARINC424.nb.html),
  SHA-256 `f0dce4a820fd14c00d41cc478ed97c33ea171d8bbde124ee356cc11e8f2a0c3b`.
  Only the fixture's HTML file is needed. Its source and referenced resources are not executed.
