# Independent status pie fixture

`generate.py` created the editable Writer chart with LibreOffice 24.2.7.2 and
`python3-uno` on Linux. It assigns green to Pass, an orange wide diagonal hatch
to Fail, and an outlined no-fill style to Unknown. It writes `status-pie.docx`,
`status-pie.odt`, and a PDF of the source Writer document. The PNG references
were rasterized at 72 DPI with Poppler `pdftoppm` 24.02.0.

| File | Role | SHA-256 |
| --- | --- | --- |
| `status-pie.docx` | Native package under test | `F1E936C2E99F2C5F7856121DCD87516FF0D54FC6D2B3F6A8F25FD0C5E8842D4E` |
| `status-pie.odt` | Producer's source package | `E2BD1D52EF351D6372171C3A451BAE59878058056A392252C6D0FADAC8118F85` |
| `status-pie-source-reference.pdf` | Source Writer document, before DOCX reopen | `61FAAB038743889F97345EA347AD2B14C663B6E82593C6540E223E82E1F160F7` |
| `status-pie-source-reference.png` | Raster of source PDF | `09DA2351B9E151D7962B84C3207955AE75BD9D0FC190C22858944D3FFEC53B54` |
| `status-pie-docx-reference.pdf` | LibreOffice PDF after reopening the DOCX | `A509F7CA788E0D8A718535A86B240BBE38EE2D04B773634E442A663356202D60` |
| `status-pie-docx-reference.png` | Raster of reopened-DOCX PDF | `FA082BD777BF9C51EFB68E9AF7E3BBFBE9A34407A297BA6A2B1FEAC03D42702F` |

The DOCX export changes the pie's direction relative to the source Writer
document. Compare OfficeIMO's DOCX rendering with `status-pie-docx-reference.png`,
not the source reference. The two references intentionally retain this difference.

To reproduce, run `generate.py` in an isolated copy of this directory using
`/usr/bin/python3`, then reopen its DOCX with LibreOffice's `--headless
--convert-to pdf` in a separate user profile. Rasterize each PDF using
`pdftoppm -f 1 -singlefile -png -r 72`. Recheck the native package and image
contents before replacing the checked-in fixtures; the hashes above identify
the exact qualified files.

## Area point paint

`generate-area.py` uses the same LibreOffice and Poppler versions. It assigns
four different point fills to one editable area series and writes those `dPt`
overrides into the DOCX. Both the source and reopened-DOCX references render
one opaque blue series area: the point fills do not create colored area
segments. The chart also has no series outline, midpoint category crossing,
and textual A–D categories with an inert date format on the category axis.

| File | Role | SHA-256 |
| --- | --- | --- |
| `area-point.docx` | Native package under test | `13B9ADF6B884607A08DB33B86D9210ABCBCEB69C8A073370903B3496B26AD0DE` |
| `area-point.odt` | Producer's source package | `5AC9DE0A2B122055EC10B4CD37A8DEA4A666EBDA387E950F9A7C4D81877BEAEA` |
| `area-point-source-reference.pdf` | Source Writer document | `0D57F1A543345FA37CE1BB2E609697D7ACEDE87040C767E83C6402E6F950D9FC` |
| `area-point-source-reference.png` | Raster of source PDF | `662343E8A014375EDDBAC305A7EC43FD1C60DBBF9093F9158F388C57F35F3E1E` |
| `area-point-docx-reference.pdf` | LibreOffice PDF after reopening the DOCX | `0C7930CC179DE6E5A748556FFABFF55C951CD82552F1BD3C8440182CD5D1499C` |
| `area-point-docx-reference.png` | Raster of reopened-DOCX PDF | `38869BE99FDEBEEAA70F8B58A9319AE74416FA6597C85DACA81EBD2B57909AC0` |

The static projection preserves the cached data and series fill, omits the
native point fills with an explicit `ChartPointStylesUnsupported` diagnostic,
and leaves the editable DOCX unchanged. The rendered area uses the opaque
series fill and no outline. This fixture does not qualify automatic numeric
tick spacing or page-level Word placement.
