# Real-world corpus evidence

OfficeIMO uses a bounded external-corpus lane to discover parser and routing failures that curated regression fixtures may not expose. The lane exercises the public `OfficeIMO.Reader.All` normalized-read contract across content-detected Word, Excel, PowerPoint, PDF, HTML, and RTF inputs.

This is discovery evidence, not a compatibility percentage. The source corpus is a convenience sample, related files may share producers or templates, and a successful normalized read does not prove visual fidelity or lossless editing.

## Evidence source and provenance

The scheduled lane uses one fixed [Govdocs1 ZIP chunk](https://digitalcorpora.org/corpora/file-corpora/files/) as a repeatable baseline. Govdocs1 contains files collected from public `.gov` web servers and explicitly warns that file extensions are only suggestions. The runner therefore requires medium- or high-confidence content evidence before assigning a format stratum; an extension alone cannot make a file eligible.

Each report records:

- the corpus identifier, source URL, and downloaded archive SHA-256;
- the runtime, operating system, limits, selection algorithm, and operation;
- every discovered file's size, extension, content hash when read, content kind and confidence, effective kind, stable detection evidence, and separate classification/read timings;
- exact eligible, duplicate, selected, completed, diagnostic, failure, timeout, and oversize counts;
- stable diagnostic codes and exception types, without extracted document content or exception messages.

Raw corpus files and generated document content are never uploaded. Source names are withheld by default. The maintained Govdocs1 workflow opts in to its public numeric corpus paths so an observation can be reproduced without rescanning every hash. Reports are retained as workflow artifacts for 30 days.

Treat every corpus file as untrusted. The archive may contain malformed content, scripts, macros, embedded objects, or other active material. The lane does not launch Office, LibreOffice, browsers, scripts, or embedded payloads.

## Why 100 files per format

The monthly target is up to 100 unique files in each available format stratum, capped at 600 selected files in total. Exact duplicates are counted but sampled once. Selection sorts unique content by SHA-256 and then takes files round-robin across strata, so the same archive and settings select the same content on every platform.

One hundred observations is a practical defect-discovery budget. If observations were independent and randomly sampled, zero failures would put the familiar approximate 95% upper bound near 3% under the rule of three. Govdocs1 does not satisfy those assumptions, so the calculation explains the budget rather than proving a reliability rate. Reports show underfilled strata instead of padding denominators or borrowing files from another format.

The fixed chunk keeps the monthly input population stable; code, runtime, and runner changes still matter when comparing results. Manual runs can select another three-digit chunk for broader discovery, but results from different chunks must not be presented as a trend without accounting for the changed population.

## Isolation and outcomes

Classification and normalized reading each run in a separate child process per file. The coordinator applies a 50 MiB input cap, a 30-second timeout per stage, a bounded traversal count, and limited parallelism. A timeout kills the worker process tree so one malformed input cannot stall the complete measurement.

The report keeps outcomes factual:

| Outcome | Meaning |
| --- | --- |
| `completed` | The normalized read returned without warning or error diagnostics. |
| `completed-with-warnings` | The read returned and OfficeIMO emitted at least one warning but no error diagnostic. |
| `completed-with-errors` | The read returned an envelope containing at least one error diagnostic. |
| `rejected-by-policy` | OfficeIMO deliberately rejected the input under an explicit permission or bounded-resource policy. |
| `failed` | The selected file caused an exception or worker failure. |
| `timed-out` | The selected file exceeded its isolated probe deadline. |
| `classification-failed` / `classification-timed-out` | The file could not be assigned to a sample stratum. |
| `skipped-oversize` | The file exceeded the byte cap before parsing. |
| `duplicate` | The same SHA-256 content was already eligible for sampling. |
| `not-selected` | The file was eligible but outside the deterministic sample budget. |
| `not-eligible` | Medium-confidence content evidence did not place the file in a requested stratum. |

The workflow is intentionally non-gating: its successful status means measurement completed, not that every selected file completed. Failures and timeouts produce workflow warnings and remain explicit in the JSON and Markdown reports. A download or archive-validation failure makes the job fail as `not-measured`; it cannot masquerade as clean evidence.

## Running the evidence lane

Use the **Real-world corpus evidence** workflow for the maintained download, provenance, extraction, reporting, and artifact-retention path. It runs monthly and accepts manual chunk, sample-size, and timeout inputs.

To measure an already extracted local corpus:

```powershell
dotnet run --project Build/RealWorldCorpus/OfficeIMO.RealWorldCorpus.Tool.csproj `
  --framework net10.0 --configuration Release -- run `
  --input C:\corpora\govdocs1-000 `
  --json artifacts\real-world-corpus\report.json `
  --markdown artifacts\real-world-corpus\report.md `
  --corpus-id govdocs1-000 `
  --source-uri https://downloads.digitalcorpora.org/corpora/files/govdocs1/zipfiles/000.zip `
  --archive-sha256 <verified-archive-sha256> `
  --max-per-format 100 `
  --max-total 600
```

Keep both report paths outside the input directory so evidence from an earlier run cannot become part of a later inventory.

The repository contract check uses only synthetic, project-owned inputs:

```powershell
pwsh -NoProfile -File Build/Test-RealWorldCorpusContract.ps1
```

## Turning discovery into regression proof

An external-corpus observation is a lead, not a permanent test by itself. Reproduce the issue, identify the owning format contract, minimize the input without removing the defect, verify redistribution terms and provenance, and add the smallest useful fixture to that owner's curated corpus. The focused test should assert the affected semantic, diagnostic, security, or round-trip contract—not merely that a historical file opens.

The open roadmap item to expand cross-producer fixture corpora remains the long-term owner for that promotion work. This lane supplies repeatable discovery and accounting; curated owner tests supply durable release evidence.
