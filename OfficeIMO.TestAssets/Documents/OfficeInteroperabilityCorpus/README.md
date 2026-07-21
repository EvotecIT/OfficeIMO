# Office Interoperability Corpus Contract

`corpus-manifest.json` is the executable inventory for the checked-in DOC, XLS,
XLSB, and PPT compatibility corpora. Each collection records its concrete
OfficeIMO format identifier, producer and provenance, supported test directions,
applicable oracle lanes, and the expected SHA-256 for every binary artifact.

The manifest test fails when a tracked binary is missing or changed, when an
untracked binary appears in a collection, or when a legacy DOC/XLS artifact no
longer has an approved import report. It also rejects unknown or duplicate
directions and oracles. Behavioral tests remain the source of truth for
projection, diagnostics, editing, conversion, and preservation; the manifest
makes sure those tests continue to run against the intended files.

The manifest distinguishes five directions: legacy import, new legacy write,
legacy round trip, modern-to-legacy conversion, and legacy-to-modern conversion.
Oracle declarations say how a collection is intended to be checked: structural,
semantic, diagnostic, visual, Microsoft Office open, LibreOffice open, or
LibreOffice render. A declaration does not pretend that every platform can run
every oracle. The portable gate enforces managed structural and semantic checks;
Windows Office and LibreOffice jobs execute their declared lanes where those
applications are available.

When adding or replacing a fixture:

1. Keep it small, focused, licensed for repository use, and free of sensitive
   data.
2. Record its producer and source in the collection's provenance file.
3. Generate or approve the legacy import report where applicable.
4. Add the artifact, its SHA-256, concrete format id, directions, and oracles to
   the manifest.
5. Add or update a focused behavioral test for the feature represented by the
   fixture.

Run the focused contract locally with:

```powershell
./Build/Test-OfficeInteroperabilityGate.ps1 -Suite Corpus
```

Run the declared LibreOffice open/convert/render lanes and retain a JSON report
plus generated artifacts with:

```powershell
./Build/Test-OfficeCorpusLibreOffice.ps1 -OutputDirectory ./artifacts/office-corpus-libreoffice
```

This lane also creates fresh OfficeIMO DOC, XLS, XLSB, and PPT outputs from
modern sources, then asks LibreOffice to open and convert all four and to render
the generated PPT. Use `-SkipOfficeImoGeneratedOutputs` only when isolating the
checked-in corpus from writer-output validation.

On a Windows machine with desktop Word, Excel, and PowerPoint installed, include
the opt-in Microsoft Office source/conversion oracle with:

```powershell
./Build/Test-OfficeInteroperabilityGate.ps1 -Suite Corpus -MicrosoftOffice
```
