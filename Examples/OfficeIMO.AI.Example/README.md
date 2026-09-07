# Document AI headless example

This .NET 10 executable loads plain text, PDFs, or raster images, runs an `OfficeIMO.AI` operation, and writes review artifacts. Reader, PDF rendering, CSV, and Excel remain owned by their existing packages. The example's dependency set is deliberately broader than the two AI packages.

Run commands from the repository root. Restore the projects using a feed containing their declared dependencies, then build the example:

```powershell
dotnet build Examples/OfficeIMO.AI.Example/OfficeIMO.AI.Example.csproj -c Release
dotnet run --project Examples/OfficeIMO.AI.Example/OfficeIMO.AI.Example.csproj -c Release --no-build -- --help
```

## Synthetic invoice

The default input is a short synthetic invoice. This command sends it to ChatGPT using the existing Codex login and writes a JSON evidence report, CSV fields, and an Excel workbook:

```powershell
dotnet run --project Examples/OfficeIMO.AI.Example/OfficeIMO.AI.Example.csproj -c Release --no-build -- --allow-remote --codex-session --output output/invoice
```

`--allow-remote` authorizes the selected source evidence to leave the machine. `--images` additionally sends selected page images. Without `--codex-session`, the native route uses IX's default authentication selection. Existing output files are never overwritten.

## Process a file

Save this request as `request.json`:

```json
{
  "operation": "ExtractFields",
  "instruction": "Extract the invoice number, total and due date.",
  "culture": "pl-PL",
  "pages": [1],
  "fields": [
    { "name": "invoiceNumber", "type": "String" },
    { "name": "total", "type": "Decimal" },
    { "name": "dueDate", "type": "Date", "dateFormat": "dd.MM.yyyy" }
  ]
}
```

```powershell
dotnet run --project Examples/OfficeIMO.AI.Example/OfficeIMO.AI.Example.csproj -c Release --no-build -- --source invoice.pdf --request request.json --allow-remote --codex-session --output output/extraction
```

Add `--images` for scanned pages or visual tables. Use `"operation": "Parse"` with an instruction and no fields to produce proposed blocks/tables and `proposed-reader.json`. `Ask`, `Explain`, and `Summarize` likewise take an instruction without field definitions.

The reader enforces its normal source permissions. The example does not request passwords or bypass protected files. Raster input metadata is excluded from text evidence; it is not treated as recognized document text. Scans may retain a partial-result status because Reader reports OCR or rendering limitations, even when a proposed table matches an evaluation fixture.

## Local or another compatible provider

The operation and request file stay the same:

```powershell
dotnet run --project Examples/OfficeIMO.AI.Example/OfficeIMO.AI.Example.csproj -c Release --no-build -- --source notes.txt --request request.json --endpoint http://127.0.0.1:11434/v1 --local --model your-installed-model --prompted-json --output output/local
```

Use a request appropriate for the file: a plain-text source may have no native page 1. A hosted compatible endpoint requires HTTPS, `--allow-remote`, and, when needed, `OFFICEIMO_AI_API_KEY` in the process environment. Do not put credentials in the endpoint URL or request JSON. Configure a vision-capable model before adding `--images`.

For LM Studio, start its server on loopback and use the loaded model's API identifier with `--endpoint http://127.0.0.1:1234/v1 --local --model <identifier>`. Use `--request-characters 16000` to exercise a smaller request profile. Character limits do not predict token counts exactly; qualify the chosen model's context and output budget. Omit `--prompted-json` when the configured runtime enforces JSON Schema. A vision-capable local model uses the same `--images` option and image evidence checks as a hosted model.

## Evaluation corpus

`--evaluate` generates the versioned synthetic corpus and runs each selected case once by default. `--repeat 2` or `--repeat 3` measures repeated-call stability. It checks English and Polish fields, exact table cells, images, scanned and rotated PDFs, mixed sources, abstention, conflicting values, source-instruction isolation, summaries, explanations, missing and ambiguous fields, refunds, regional columns, a directed-flow diagram and long-document synthesis. It has a one-hour run deadline and uses each operation's own request limits. It never reads arbitrary input documents in evaluation mode.

```powershell
dotnet run --project Examples/OfficeIMO.AI.Example/OfficeIMO.AI.Example.csproj -c Release --no-build -- --evaluate --allow-remote --codex-session --output output/evaluation
```

Each case saves its source, hash-bound report, provider response, and applicable Reader/CSV/Excel artifacts. `evaluation.json` records the profile, corpus version, exact assertions, statuses, omissions, usage when available, and elapsed times. Fonts come from the host's embeddable system fonts; preserve the generated source files and hashes when comparing runs on different machines. `--case contradictory` selects one case for diagnosis and produces a report for that subset only.

Use `--split development` while refining behavior and `--split heldout` for separately declared challenge values and layouts. Once a held-out failure guides a change, treat that case as regression evidence and use fresh reserve cases for the next independent check. Expected values are declared in the corpus before inference, not inferred from model output.

Every case has typed gold values declared before inference. Reports include exact field matches, position-sensitive table-cell precision/recall, fact-marker recall, request attempts, maximum measured request size, synthesis state, and text coverage. Extra returned fields or tables fail the corresponding exact-match check. Memory samples describe the evaluation process only; they exclude model-server and GPU memory. A case failure does not discard earlier results, and reports are checkpointed after each completed case. `--request-characters` sets a profile-specific request bound for smaller contexts.

For an installed Copilot CLI, use `--copilot --text-only --allow-remote --model <available-model>`. `OFFICEIMO_AI_API_KEY` can provide its accepted GitHub token. `--text-only` explicitly excludes image-required cases and records their identifiers in the report. The SDK disables tools and ambient configuration for this route. Model availability and subscription access must be verified for the account.

Every selected repetition must meet its field/table/status assertion for the run to pass. Claim checks are smoke checks; they do not replace independent semantic assessment. A passing finite synthetic corpus is not a general document-accuracy guarantee. See the [support matrix](../../Docs/officeimo.document-assistant-design.md) for unverified deployment and quality coverage.

Exit codes: `0` completed or evaluation assertions passed; `1` partial, insufficient, invalid, or a failed evaluation assertion; `2` setup/input failure; `3` cancellation or timeout. The general file-processing path reports diagnostic codes and exception types without raw provider errors. Evaluation response recording is limited to the synthetic corpus.
