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
