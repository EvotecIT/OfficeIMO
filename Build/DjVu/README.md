# DjVu fixture and reference validation

Normal tests consume the fixed fixtures in `OfficeIMO.DjVu.Tests/Fixtures`. Their manifest records producer roles, ownership, sizes and SHA-256 hashes. Native encoders and decoders are opt-in validation tools; product packages, normal tests and ordinary builds do not invoke or download them.

To independently produce the authored colour, grayscale, background-sampling, palette, rotated text, outline and compressed-byte cases, supply installed DjVuLibre tools and a new scratch directory:

```sh
python3 Build/DjVu/generate_authored_fixtures.py \
  --djvulibre-bin /path/to/djvulibre/bin \
  --output /path/to/task-scratch/djvu-fixtures
```

Optional `--cjpeg /path/to/cjpeg` adds a managed-JPEG comparison case and requires a JPEG-enabled DjVuLibre build. Optional `--minidjvu /path/to/minidjvu` adds losslessly encoded shared JB2 dictionaries in bundled and indirect documents. The script consumes checked-in authored PBM and raw BZZ inputs, and writes source images, reference output and a new hash manifest. It does not replace repository fixtures. Native encoder versions can change compressed bytes; compare decoded content as well as file identities.

Optional `--jb2-work-producer /path/to/producer` adds repeated comments, zero-area symbols and fully clipped repeated masks. Build `generate_jb2_work_fixtures.cpp` against the same validation-only DjVuLibre source headers, configured `config.h` and native library. For example, with those paths in shell variables:

```sh
c++ -std=c++14 -DHAVE_CONFIG_H -I"$DJVU_SOURCE/libdjvu" -I"$DJVU_BUILD" \
  Build/DjVu/generate_jb2_work_fixtures.cpp -L"$DJVU_PREFIX/lib" -ldjvulibre \
  -Wl,-rpath,"$DJVU_PREFIX/lib" -o "$TASK_SCRATCH/jb2-work-producer"
```

The producer uses native ZP arithmetic encoding for authored record sequences and independently decodes their dimensions, symbols and placements with DjVuLibre. Its large masked page qualifies the empty column used by the regression test; it is not an exact-raster claim for the entire page. Fixed tests enforce comment limits across records and shared decoder budgets, and preserve empty-symbol geometry and clipped-region pixels. They do not impose host timing thresholds.

The progressive gradient and constructed MMR fixtures remain fixed inputs. MMR reference pixels come from DjVuLibre decoding; they are not proof of an independent MMR encoder. Archive fixtures retain original page FORM bytes from an identified public-domain source; the archive manifest records extraction offsets and source identity. Full archival books and specification documents are supplied separately to the [all-page comparison runner](../../OfficeIMO.DjVu.Verification/README.md).

The [dated summary](Evidence/2026-10-09/rendering-summary.json) records complete native comparisons and the short-edge case outside the declared acceptance profile. It is qualification evidence for those source bytes, not a format-wide correctness claim.
