# OfficeIMO.AI

`OfficeIMO.AI` provides read-only document questions, explanations, summaries, field extraction, and proposed document structure for .NET 10. It accepts an immutable Reader snapshot and a caller-supplied `IOfficeAiExecutor`. The package depends on `OfficeIMO.Reader.Core`; format readers, rendering, OCR, and model clients are selected by the host.

Use [OfficeIMO.AI.IntelligenceX](../OfficeIMO.AI.IntelligenceX/README.md) for ChatGPT, native Copilot, or an OpenAI-compatible endpoint. The [headless example](../Examples/OfficeIMO.AI.Example/README.md) loads PDFs, text, and images and writes JSON, CSV, and Excel review artifacts.

## Add to a .NET 10 application

To build from a source checkout, create an application beside the `OfficeIMO` directory and reference the engine project:

```shell
dotnet new console --framework net10.0 --name DocumentAssistant
dotnet add DocumentAssistant/DocumentAssistant.csproj reference OfficeIMO/OfficeIMO.AI/OfficeIMO.AI.csproj
dotnet build DocumentAssistant/DocumentAssistant.csproj
```

The engine reference brings in Reader Core. Supply your own `IOfficeAiExecutor`, or add the [IntelligenceX adapter](../OfficeIMO.AI.IntelligenceX/README.md#add-the-adapter-from-source).

## Extract named fields

This method reads a plain-text invoice. Register the relevant Reader adapter for other formats.

```csharp
using OfficeIMO.AI;
using OfficeIMO.Reader;

static async Task<OfficeAiResult> ExtractAsync(
    IOfficeAiExecutor executor, Stream source, bool allowRemote,
    CancellationToken cancellationToken = default) {
    using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
    deadline.CancelAfter(TimeSpan.FromMinutes(3));
    var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
    var document = await OfficeAiDocument.ReadAsync(reader, source, "invoice.txt",
        cancellationToken: deadline.Token);
    return await new OfficeAiEngine(executor).RunAsync(document, new OfficeAiRequest {
        Operation = OfficeAiOperation.ExtractFields,
        Instruction = "Extract the invoice total and due date.",
        Culture = "en-US",
        AllowRemoteProcessing = allowRemote,
        Fields = new[] {
            new OfficeAiFieldDefinition("total", OfficeAiFieldType.Decimal),
            new OfficeAiFieldDefinition("dueDate", OfficeAiFieldType.Date, "yyyy-MM-dd")
        }
    }, cancellationToken: deadline.Token);
}
```

`AllowRemoteProcessing` defaults to `false`. The host must obtain a deliberate choice before sending evidence to a profile with `IsLocal = false`. Reusing that choice within the authorized scope does not require another prompt per batch.

## Operations and evidence

| Operation | Output |
| --- | --- |
| `Ask` | Claims answering the question, each with source references |
| `Explain` | Source-linked explanations within the selected scope |
| `Summarize` | Source-linked summary combining validated batch drafts within the operation budget |
| `ExtractFields` | Exactly the requested field names, with raw values, invariant normalized values, status, and references |
| `Parse` | Proposed Reader blocks and rectangular tables, with references |

Select one-based `Pages`, `EvidenceIds`, or both. Unknown identifiers and selections outside the page/image scope are rejected before execution. Empty selections mean all captured evidence. `IncludeImages` must be explicit, and the selected execution profile must support vision.

For follow-up questions, `ConversationContext` accepts up to 8000 characters of prior discussion. It is measured in the request budget and supplied separately as untrusted context. Previous answers cannot substitute for current source evidence or become citations. Hosts must clear or rebind that context when the document, scope or account changes; Studio keeps at most three exchanges for the current snapshot.

`Ask` and `Explain` evaluate each batch independently. When selected evidence spans multiple batches, validated batch claims remain available, but the result is `Partial` with `cross-batch-reasoning-not-supported`. The engine cannot determine an answer that requires relating facts across those batches. A single-batch request does not have this limitation.

Table evidence includes the title and column headers even when there are no rows. Each row retains its title and column labels, and `sourceBlockId` identifies its table and row in execution requests. When Reader supplies a shared anchor for a table placeholder and its data, `sourceAnchor` preserves that association across narrative blocks, table headers and rows, including untitled tables beneath section headings.

Each snapshot retains the SHA-256 of the original bytes, Reader page provenance, source block identifiers, and available source geometry. A separate `SnapshotHash` binds results and exports to the exact evidence projection, image payload hashes and coverage state; the same original bytes with different observations are not interchangeable. `FromReadResult` is a trusted-adapter entry point: its caller must enforce source permissions and ensure the supplied Reader result and images describe those exact bytes. `OfficeAiImage` takes verified dimensions from the rendering/image owner and copies its payload. It does not decode or certify an image itself.

Text references must contain an exact contiguous source quote. `QuoteStart` records its zero-based UTF-16 offset within the original evidence record. Unknown references and altered quotes invalidate the response. Image references have no text-match claim. A matching quote proves where text occurs; it does **not** establish that the model's interpretation follows from it. Every result has `RequiresReview = true`.

## Result states

`Completed` means all selected evidence reached requests whose responses satisfied the structural contract. It does not mean every claim is correct or every visible detail was recognized. `Partial` identifies omissions, source warnings (including truncated tables and chunk warnings), pages without evidence, or invalid scalar normalization. `InsufficientEvidence` identifies a valid abstention. `InvalidResponse` means no batch produced a validated result after provider or response-validation failures; diagnostic codes distinguish those failures without exposing raw errors.

The response schema matches the selected operation. For field extraction, the executor returns a `fields` object with every requested field key (`field1`, `field2`, and so on) required; each value contains `status`, `rawValue` and `evidence`. Request metadata pairs each stable key with the original field name, so punctuation and Unicode in caller names do not become schema-key restrictions. The schema rejects omitted or unrequested keys and constrains the value and evidence shape for present, missing and uncertain fields. Local validation also rejects duplicate JSON keys. The public `OfficeAiResult.Fields` collection retains the original names and requested field order. Field extraction requires empty claims, blocks and tables; parsing requires empty claims and fields; questions, explanations and summaries require empty fields, blocks and tables. The same rules are checked locally for every provider.

Field states distinguish `Present`, `Missing`, `Ambiguous`, `Conflicting`, `Invalid`, and `NotEvaluated`. A field is `Missing` only when the processed evidence did not provide it; incomplete source coverage uses `NotEvaluated` for otherwise missing fields. Conflicting values are not collapsed into one normalized value. Decimal normalization uses the explicit culture and validates grouping. Dates require an exact format. Integer and Boolean normalization accept their ordinary signed-integer and `true`/`false` forms.

## Budgets and cancellation

`OfficeAiLimits` bounds captured bytes, retained observations, pages, request text, image payloads/pixels, response text, result sizes, request count, and duration. `MaxDocumentImages` bounds image count independently of `MaxPages`, so several images can describe one page. Generation schemas use the same `MaxResultItems`, `MaxTableCells`, and `MaxTableColumns` bounds as local response validation, with equal column and row widths. `MaxTableColumns` defaults to 32 and supports up to 100; wider bounds need more schema context. For smaller model contexts, reduce it to the widest table the operation needs. Format readers and renderers also need their own allocation and decoding limits. Snapshot limits do not replace those owners' parser limits.

The engine includes the executor's prompt-wrapper measurement when batching. Oversized text records are split into contiguous windows at nearby natural boundaries without splitting a UTF-16 surrogate pair. The snapshot stays unchanged, and validated citations map back to its original identifiers and offsets. `ProcessedTextRanges` records successful windows. `ProcessedEvidenceIds` contains fully processed records; a record with any unprocessed text remains in `OmittedEvidenceIds`.

Multi-batch summaries combine validated drafts through bounded reduction passes. Each combined claim references known draft identifiers; the engine attaches their original citations and rejects unknown identifiers or omitted draft groups. This preserves reference lineage, not a proof of semantic entailment. `SynthesisStatus` reports whether combination completed. If the request budget, response validation, or pass limit prevents completion, validated drafts remain available with `Partial` and `summary-synthesis-incomplete`. `RequestCount` includes map and synthesis attempts. `MaxSynthesisPasses` defaults to three, and every pass shares `MaxRequests` and the operation deadline.

Use one linked cancellation token for read, render/OCR, inference, and artifact writing when the host needs one end-to-end deadline. Cancellation stops waiting and discards late results. An executor that ignores cancellation retains its execution gate until its actual work settles, preventing overlapping calls through that executor. Callers remain responsible for the lifetime of a supplied stream or provider that continues working after cancellation. The engine makes no automatic repair request.

## Review artifacts

`OfficeAiArtifacts.SerializeReport(document, result)` writes a versioned JSON report with source evidence, image descriptors, coverage, references, and result status. It omits encoded images and connection credentials. Saving the report may still save private source text; choose the output location accordingly.

For `Parse`, `CreateProposedReadResult` produces Reader's canonical transport model with an explicit AI-proposal warning. The example exports tables through `OfficeIMO.CSV` and `OfficeIMO.Excel`, preserving values as text and using CSV formula-injection protection. It reopens generated Reader JSON and Excel output. No engine operation modifies the source file or applies a proposed edit.

## Executor contract

Implement `IOfficeAiExecutor` to use another model client. Supply an immutable profile, report whether the actual route is local and whether it accepts images/enforces schemas, measure transport prompt text in `MeasureRequestCharacters`, and return one bounded response from `ExecuteAsync`. A truncated generation must set `IsComplete = false`. Leave unavailable usage counters null. The provider boundary must not give document content access to tools, files, or arbitrary network actions.

Profile capability declarations require independent qualification. The engine applies the same local response checks to schema-enforced and prompted-JSON output; the latter reports its weaker generation guarantee. See the [architecture and support matrix](../Docs/officeimo.document-assistant-design.md) for current coverage and limits.
