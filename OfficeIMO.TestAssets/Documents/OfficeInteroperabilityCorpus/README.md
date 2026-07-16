# Office Interoperability Corpus Contract

`corpus-manifest.json` is the executable inventory for the checked-in DOC, XLS,
and XLSB compatibility corpora. Each collection records who produced the files,
where their provenance is documented, which behavior they exercise, and the
expected SHA-256 for every binary artifact.

The manifest test fails when a tracked binary is missing or changed, when an
untracked binary appears in a collection, or when a legacy DOC/XLS artifact no
longer has an approved import report. Behavioral tests remain the source of
truth for projection, diagnostics, editing, conversion, and preservation; the
manifest makes sure those tests continue to run against the intended files.

When adding or replacing a fixture:

1. Keep it small, focused, licensed for repository use, and free of sensitive
   data.
2. Record its producer and source in the collection's provenance file.
3. Generate or approve the legacy import report where applicable.
4. Add the artifact and its SHA-256 to the manifest.
5. Add or update a focused behavioral test for the feature represented by the
   fixture.

Run the focused contract locally with:

```powershell
./Build/Test-OfficeInteroperabilityGate.ps1 -Suite Corpus
```
