# PDF accessibility check support matrix

`PdfDocumentProof.CheckAccessibility()` (PowerShell: `Test-OfficePdfAccessibility`) checks any loaded PDF against
PDF/UA-1 (ISO 14289-1). It reports every Matterhorn Protocol 1.1 failure condition plus the five ISO 14289-1
requirements that veraPDF checks without a Matterhorn number. `PdfAccessibilityConditions.All` lists the catalogue
with each condition's coverage, the veraPDF PDF/UA-1 rules that test the same requirement, and the PDF/UA-2 rules
where the requirement carries over.

The result never certifies conformance. Conditions that need a person's judgement are always `NeedsReview`,
never `Passed`.

## Coverage

| Checkpoint | Automatic | Manual review | Not checked yet |
| --- | --- | --- | --- |
| 01 Real content tagged | 01-003, 01-004, 01-005, 01-007, plus MarkInfo Marked (6.2-1), StructTreeRoot (7.1-11) and the structure element P entry (7.1-12) | 01-001, 01-002, 01-006 | |
| 02 Role mapping | 02-001, 02-003, 02-004 | 02-002 | |
| 03-05 Flickering, colour, sound | | all | |
| 06 Metadata | 06-001, 06-002, 06-003 | 06-004 | |
| 07 Dictionary | 07-001, 07-002 | | |
| 08 OCR | | all | |
| 09 Appropriate tags | 09-004 to 09-008 | 09-001 to 09-003 | |
| 10 Character mappings | | | 10-001 (glyph level; see 31-027) |
| 11 Natural language | 11-001 to 11-006, plus valid Lang values (7.2-29) | 11-007 | |
| 12 Stretched characters | | 12-001 | |
| 13 Graphics | 13-004 | 13-001 to 13-003, 13-005 to 13-008 | |
| 14 Headings | 14-002, 14-003, 14-006, 14-007 | 14-001, 14-004, 14-005 | |
| 15 Tables | 15-003 | 15-001, 15-002, 15-004, 15-005 | |
| 16 Lists | | all | |
| 17 Mathematical expressions | 17-002 | 17-001 | 17-003 |
| 18 Page headers and footers | | all | |
| 19 Notes and references | 19-003, 19-004 | 19-001, 19-002 | |
| 20 Optional content | 20-001 to 20-003 | | |
| 21 Embedded files | 21-001 | | |
| 22-24, 27, 29 | | all (23-001 and 27-001 have no specific test) | |
| 25 XFA | 25-001 | | |
| 26 Security | 26-001, 26-002 | | |
| 28 Annotations | 28-002, 28-004 to 28-012, 28-014 to 28-017, plus single-widget Form elements (7.18.4-2) | 28-001, 28-003, 28-013 | 28-018 |
| 30 XObjects | 30-001, 30-002 | | |
| 31 Fonts | 31-001 to 31-009, 31-019 to 31-021, 31-024, 31-027 | 31-010 | 31-011 to 31-018, 31-022, 31-023, 31-025, 31-026, 31-028 to 31-030 |

## Behaviour and limits

- Fonts are checked when page content selects them with `Tf` and shows text; fonts used only in annotation
  appearances or Type 3 glyph procedures are not checked. 31-009 skips fonts used only with rendering mode 3.
- 31-027 accepts simple fonts with a standard encoding, non-symbolic TrueType fonts, unembedded standard 14 Type 1
  fonts and Type 0 fonts with an Adobe GB1, CNS1, Japan1 or Korea1 collection. Type 1 and Type 3 fonts that rely on
  glyph names are `Indeterminate` until the Adobe Glyph List check exists.
- 31-003 follows veraPDF: it fails when the CIDFont Supplement is greater than the CMap Supplement. With a
  predefined CMap the CMap's Supplement isn't carried, so a CIDFont Supplement above 0 is `Indeterminate`.
- 06-002 accepts `pdfuaid:part` 1 or 2, so PDF/UA-2 files keep their identification; veraPDF's PDF/UA-1 profile
  requires 1.
- 15-003 treats a table as organised with Headers and IDs only when every TD names headers that are TH cells of the
  same table. veraPDF resolves each cell's headers instead; the two agree on the veraPDF corpus.
- 09-004 to 09-008 also fail a grouping element (Table, TR, THead, TBody, TFoot, L, TOC, Ruby, Warichu) that
  holds content directly. veraPDF checks only child elements, so these findings carry no veraPDF rule.
- Form XObjects inherit the font and rendering mode in effect at `Do`; fonts selected through an ExtGState `Font`
  entry are checked like fonts selected with `Tf`.
- Table regularity (veraPDF 7.2-15 and 7.2-41 to 7.2-43) is not checked.
- When the document permissions deny both copying and accessibility extraction, only the catalog
  MarkInfo, ViewerPreferences, catalog Lang and encryption conditions run; the rest are `Indeterminate`.
- At most 100 findings are kept per condition; `OccurrenceCount` keeps the total.
