# DjVu independent raster verification

This opt-in runner compares every page against an explicitly supplied DjVuLibre `ddjvu` executable. It is outside the normal solution, is not packable, and is not a product dependency or automatic download route.

```sh
dotnet run -c Release --project OfficeIMO.DjVu.Verification -- \
  /path/to/book.djvu /path/to/ddjvu /path/to/report.json
```

Each page is rendered at native DPI with display rotation. The runner streams reference RGB rows, validates dimensions and complete output, and records both pixel hashes, differing samples, maximum and mean channel differences, source hash, source length, DPI, text status, and reference version. Reports are checkpointed every eight pages. Exit code zero requires all pages to satisfy the declared maximum-four/mean-0.15 profile; parsing, unsupported-codec, and comparison failures remain visible per page. Exact hash equality is reported separately.

The [support contract](../OfficeIMO.DjVu/SUPPORT.md) identifies qualified corpora and the small-image profile outside that comparison bound. Run Release builds for corpus comparisons; host timings are not correctness gates. Supply only trusted reference tools and keep generated images, corpora, and reports in a task-owned scratch directory. The runner neither bundles nor licenses a reference executable for downstream users.
