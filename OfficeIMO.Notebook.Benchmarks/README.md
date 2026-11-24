# Notebook benchmarks

Run the optional [PowerForge notebook suite](../Build/Benchmarks/Run-NotebookBenchmarks.ps1)
to measure notebook reading, source projection, round-trip preservation and
static export. The workloads consume saved output and run without a kernel.
The benchmark project stays outside the product solution and packages.

```powershell
dotnet build OfficeIMO.Notebook.Benchmarks/OfficeIMO.Notebook.Benchmarks.csproj -c Release
./Build/Benchmarks/Run-NotebookBenchmarks.ps1 -Plan
./Build/Benchmarks/Run-NotebookBenchmarks.ps1 -Engine OfficeIMO -Cells 12 -WarmupCount 1 -IterationCount 1
```

Use a fresh PowerShell 7 session with PSPublishModule 3.0.190 or newer. The
`-ModulePath` option also accepts a source-built module. The project uses .NET 10.
Pass `-OutputRoot` to select a dedicated results directory. Keep Python and .NET
versions, environment settings, input fingerprints, and CPU placement fixed
across repeated runs.

## Reference engines

Build with `-p:EnableNotebookCompetitors=true` to include the isolated MIT-licensed
`Microsoft.DotNet.Interactive.Documents` 1.0.0-beta.26120.1 and `Verso` 1.2.3
NuGet comparisons. Their exact source revisions appear in the run metadata.
The pinned packages are restored only for that project and explicit build option.
The ordinary project build remains sufficient for an OfficeIMO-only run.

Python 3.10+ is used for input generation and verification. Use a dedicated
Python environment for reference runs and point `-PythonPath` at its interpreter.

```powershell
python3 -m venv ./Ignore/Benchmarks/notebook-venv
$referencePython = if ($IsWindows) {
    './Ignore/Benchmarks/notebook-venv/Scripts/python.exe'
} else {
    './Ignore/Benchmarks/notebook-venv/bin/python'
}
& $referencePython -m pip install -r Build/Benchmarks/notebook-reference-requirements.txt
./Build/Benchmarks/Run-NotebookBenchmarks.ps1 -Family Projection -Cells 12 -PythonPath $referencePython
```

The Reference engine uses nbformat 5.11.1 and nbconvert 7.17.2. Install Chromium
revision 1243 through the shared browser tooling or select its existing
Playwright cache for the WebPDF lane. The suite probes Chromium launch and
disables browser downloads. PDF verification requires Poppler's `pdftotext`
and `pdfimages` commands. Unavailable optional lanes appear as skipped. Check
the environment metadata for Python package versions and Chromium probe results.

## Measurement boundaries

| Family | What the operation does | Proof after timing |
| --- | --- | --- |
| `Codec` | Constructs the public notebook model or loads and serializes the whole notebook | Compares saved cell order, types, source, IDs, metadata, outputs and MIME alternatives, allowing string versus line-array serialization |
| `Projection` | Loads the same UTF-8 input and reads every ordered cell type and source string | Compares every projected cell to the source JSON. Microsoft and Verso include the UTF-8 decoding required by their string APIs |
| `Export` | Projects a prepared notebook to Markdown or HTML, or reads the input and writes complete PDF bytes | Checks each content marker, verifies PNG data in HTML/Markdown, and checks searchable text, page boxes and image counts in PDF via Poppler |
| `Fidelity` | Probes serialization against independently produced, pinned nbformat fixtures | Records per-field differences in `fidelity.json`. This family supplies compatibility evidence and has no speed comparison table |

`Codec/Parse` measures public model construction. OfficeIMO can defer decoding
some fields until access, while nbformat prepares text during model construction.
`Projection` measures materialization explicitly. The Python parse lane uses
`nbformat.reader.reads`. Public `nbformat.reads` and `writes` add schema validation
to its round trip. These API costs remain visible in the measurements.

`Export` uses `nbconvert` Markdown and lab HTML exporters, and Chromium WebPDF
with disabled cell execution and external resources. Chromium startup remains part
of each WebPDF operation. Export comparisons use synthetic inputs only. PDF text
and image checks establish content coverage. Inspect the retained files to compare
pagination, styling, math, and graphics fidelity.

`-Cells`, `-CodeLineCharacters`, `-OutputCharacters` and `-RichImagePixels`
control synthetic size. `-Case` can select `Common-12` or `Rich-12`.
`Fidelity` uses the [Jupyter fixture corpus](../OfficeIMO.TestAssets/Notebook/Jupyter/README.md).
An input a competitor rejects stays failed in the fidelity report. The
versions, sizes and exact input hashes are recorded with the samples.

```powershell
./Build/Benchmarks/Run-NotebookBenchmarks.ps1 -Family Export -Cells 12,120 -PythonPath $referencePython
./Build/Benchmarks/Run-NotebookBenchmarks.ps1 -Family Export -Operation Pdf -Cells 12 -Cold -PythonPath $referencePython
./Build/Benchmarks/Run-NotebookBenchmarks.ps1 -Family Export -Operation Pdf -Cells 12 -CodeLineCharacters 600 -PythonPath $referencePython
```

`-Cold` starts a fresh worker for every operation with zero warmups. Its time
includes process startup, imports, input loading and the first call. Cold
exports also write the output file and any associated assets before timing ends.
The cold route inherits the host's disk cache. Compare it separately from the
warmed worker and retain clipping or missing content as a failed sample.

On Windows, use `-ProcessorAffinityMask` and `-ProcessPriority` to request a
fixed processor domain. PowerForge records placement for its host. Confirm
that child workers inherit the intended settings before publishing comparisons.

PowerForge owns warmups, ordering, elapsed time and sample artifacts. Warm runs
prepare input and start workers outside the measured operation. All runs validate
artifacts afterward. Elapsed samples include worker dispatch and response, which is
visible on small cases. `WorkerManagedAllocatedBytes` reports allocation during
the .NET owner call. PowerForge's own allocation and memory metrics observe the
PowerShell runner. Comparable peak resident memory for the workers and their
browser descendants remains unqualified. Keep failed validations and examine
their artifacts before comparing rates.
