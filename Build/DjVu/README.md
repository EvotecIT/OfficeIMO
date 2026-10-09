# DjVu fixture and reference validation

Normal tests consume the fixed fixtures in `OfficeIMO.DjVu.Tests/Fixtures`. Their manifest records producer roles, ownership, sizes and SHA-256 hashes. Native encoders and decoders are opt-in validation tools; product packages, normal tests and ordinary builds do not invoke or download them.

To independently produce the authored colour, grayscale, background-sampling, palette, rotated text, outline and compressed-byte cases, supply installed DjVuLibre tools and a new scratch directory:

```sh
python3 Build/DjVu/generate_authored_fixtures.py \
  --djvulibre-bin /path/to/djvulibre/bin \
  --output /path/to/task-scratch/djvu-fixtures
```

Optional `--cjpeg /path/to/cjpeg` adds a managed-JPEG comparison case and requires a JPEG-enabled DjVuLibre build. Optional `--minidjvu /path/to/minidjvu` adds losslessly encoded shared JB2 dictionaries in bundled and indirect documents. The script consumes checked-in authored PBM and raw BZZ inputs, and writes source images, reference output and a new hash manifest. It does not replace repository fixtures. Native encoder versions can change compressed bytes; compare decoded content as well as file identities.

The progressive gradient and constructed MMR fixtures remain fixed inputs. MMR reference pixels come from DjVuLibre decoding; they are not proof of an independent MMR encoder. Archive fixtures retain original page FORM bytes from an identified public-domain source; the archive manifest records extraction offsets and source identity. Full archival books and specification documents are supplied separately to the [all-page comparison runner](../../OfficeIMO.DjVu.Verification/README.md).

The [dated summary](Evidence/2026-10-09/rendering-summary.json) records complete native comparisons and the short-edge case outside the declared acceptance profile. It is qualification evidence for those source bytes, not a format-wide correctness claim.
