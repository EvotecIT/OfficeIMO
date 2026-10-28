# Ruled tables and vector figures

`word-business-summary.pdf` is the frozen OfficeIMO PDF export of the repository's business-delivery-summary Word fixture at `c000d5ce86`. Its second page has a ruled Worksheet/Purpose table that was mistaken for a vector figure, hiding its editable text behind a picture. The manifest records its exact bytes. The tests also use the existing independent Keynote table export under `OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports`.

These are test-only inputs. The tests qualify figure classification and retained editable text, separately from complete ruled-grid reconstruction.
