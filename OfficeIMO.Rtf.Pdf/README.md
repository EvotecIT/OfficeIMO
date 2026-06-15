# OfficeIMO.Rtf.Pdf

Dependency-free RTF to PDF conversion for OfficeIMO.

This package converts the semantic `OfficeIMO.Rtf` document model into the first-party `OfficeIMO.Pdf` document model. The RTF engine remains the lossless parse/edit/write layer; PDF export is a visual/content conversion to a fixed-layout format.

Supported export coverage includes semantic paragraphs, paragraph indentation/spacing/line-height/pagination controls, section-owned blocks, section page breaks, page-starting section page setup, document and section page-border visual export, rich runs, list markers, document page setup, metadata, tables, PNG/JPEG images, bookmarks, field result text, hidden text control, footnote/endnote/annotation bodies, and running header/footer text including first-page and even-page variants. RTF can model separate borders per page side; PDF export maps the first styled RTF page border to the first-party PDF engine's uniform page border decoration.
