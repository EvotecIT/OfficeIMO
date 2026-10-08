# Missing-text raster regressions

These four PNGs were rendered from OfficeIMO's Markdown visual-baseline samples
with Poppler 26.05 using a Fontconfig setup that selected macOS Helvetica.ttc.
Most normal text disappeared although the source PDFs retained extractable text.
The images passed the existing whole-page pixel and perceptual error budgets.

`PdfDocumentRasterVisualBaselineContentTests` compares them with the corresponding
committed visual baselines to protect against accepting catastrophic visible ink
loss. The tests read the recorded images and require no native renderer or fonts.
The fixtures are generated from the repository's own samples and are test-only.
