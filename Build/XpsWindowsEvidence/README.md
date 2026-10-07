# Windows XPS rendering evidence

This opt-in runner uses Microsoft's WPF XPS package reader and offscreen renderer
to compare Microsoft XPS inputs with OfficeIMO's managed drawing and PDF
projections. It runs on Windows, stays outside the solution and does not add a
desktop framework or native codec to shipped OfficeIMO packages.

From the repository root, using the SDK selected by `global.json`:

```powershell
$repository = (git rev-parse --show-toplevel).Trim()
$sourceCommit = (git rev-parse HEAD).Trim()
$output = Join-Path $env:TEMP 'officeimo-xps-windows-evidence'
dotnet run --project Build/XpsWindowsEvidence/OfficeIMO.XpsWindowsEvidence.csproj -c Release -- $repository $output $sourceCommit
```

Use an explicit task-owned output directory. The runner writes source packages,
native/managed/PDF PNGs, converted PDFs/SVGs and `report.json`. It overwrites its
known filenames when repeated. Keep representative evidence and remove obsolete
output after inspection. Record a dirty source context when running with
uncommitted code; the report also hashes the actual loaded OfficeIMO assemblies.

The authored corpus covers stroke caps, joins and degeneracy, radial Pad/Repeat/
Reflect fields and transforms, and the existing integer JPEG/TIFF default-image
fixtures. Gradients contain the required `MappingMode="Absolute"` attribute.
An independent WPF serializer produces a two-document, three-page package using
the repository's licensed Carlito font. For each independent page the runner
checks extraction and searchable PDF text, identical WPF pixels after an
unchanged OfficeIMO save, and a visible loaded path edit after save/reopen.

`report.json` retains every case, including disagreements. RGB comparisons use
white backgrounds at 96 dpi. The stable-interior metric excludes the one-pixel
border and native 3-by-3 neighborhoods where any neighbor differs from the center
by more than 4/255 in any channel. Full-image mean/max errors and pixels above
4/255 remain separate measurements.
Inspect the renders as well as those numbers, especially high-frequency fields
and antialiased strokes. A zero exit code means evidence collection and lifecycle
checks completed; it does not mean that every rendering is equivalent. Native
open/render failures and lifecycle failures produce a nonzero exit code.

The [XPS support matrix](../../OfficeIMO.Xps/SUPPORT.md#bounded-windows-wpf-comparison)
describes the qualified scope and observed differences. This runner does not
qualify OpenXPS, interleaved packages, StoryFragments, printing or an independent
PDF renderer. The [recorded run](Evidence/2026-10-07/report.json) preserves the
report, independent WPF package and representative input/render triples. Those
triples cover a WPF page, a boundary Repeat field, coincident stroke endpoints
and an unspecified TIFF extra sample. Full generated artifacts remain in the
selected output directory.
