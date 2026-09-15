# H0 representative static-page bundle

This frozen fixture is a hand-authored static business report used to establish the OfficeIMO HTML qualification baseline. It is independent of the OfficeIMO HTML generators and exercises a normal page/resource boundary:

- document-relative external CSS
- an external OpenType font with its redistribution license
- an external SVG image
- responsive wide and narrow layouts
- paged print rules
- headings, links, lists, cards, and a data table

`manifest.json` freezes every input by byte length and SHA-256. The qualification runner rejects missing, extra, or changed files before parsing or rendering. Generated evidence belongs in an ignored artifact directory and is not part of this source bundle.

`fonts/SourceSansPro-Regular.otf` is copied from the pinned SixLabors.Fonts test fixture already retained by OfficeIMO. Source Sans Pro is licensed under the SIL Open Font License 1.1; the complete license and font-specific copyright and Reserved Font Name notice are included beside the font.
