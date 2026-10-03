using System.Globalization;
using OfficeIMO.Ocr;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrExecutionExtensions {
    /// <summary>
    /// Executes OCR over a document and its nested results with one shared budget. Newly recognized
    /// nested text is projected into each parent for Markdown, chunk and AI consumers. Existing native
    /// content and rich nested results are preserved. Call after reading a container; Reader processors
    /// already visit individual documents and must not invoke this tree operation for every node.
    /// </summary>
    public static async Task<OfficeDocumentOcrExecutionResult> ApplyOcrTreeAsync(
        this OfficeDocumentReadResult document, IOcrEngine engine,
        OfficeDocumentOcrExecutionOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (engine == null) throw new ArgumentNullException(nameof(engine));
        ExecutionOptionsSnapshot effective = ExecutionOptionsSnapshot.Create(options);
        OcrEngineExecution execution = OcrEngineRunner.CreateExecution(engine);
        var budget = new ExecutionBudget(effective);
        var timedOut = new TimedOutOcrOperationTracker();
        var active = new HashSet<OfficeDocumentReadResult>();
        var recognitions = new List<OfficeDocumentOcrRecognition>();
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        var report = new OfficeDocumentOcrExecutionReport { EngineId = execution.Id };

        async Task<OfficeDocumentReadResult> Visit(OfficeDocumentReadResult input, string id, string? path, string originalVirtualPath, int depth) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!active.Add(input)) throw new ArgumentException("Nested OCR requires an acyclic document tree.", nameof(document));
            try {
                OfficeDocumentOcrExecutionResult own = await ApplyOcrDocumentAsync(input, execution, effective,
                    budget, timedOut, id, path, cancellationToken).ConfigureAwait(false);
                foreach (OfficeDocumentDiagnostic diagnostic in own.Diagnostics) {
                    diagnostic.Location = ProjectLocation(diagnostic.Location, input.Source?.Path, originalVirtualPath,
                        path, diagnostic.Location?.BlockAnchor ?? string.Empty);
                    var attributes = diagnostic.Attributes.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
                    attributes["documentId"] = id;
                    diagnostic.Attributes = attributes;
                }
                AddReport(report, own.Report);
                recognitions.AddRange(own.Recognitions);
                diagnostics.AddRange(own.Diagnostics);
                OfficeDocumentReadResult output = own.Document;
                var nestedResults = new List<OfficeDocumentNestedResult>();
                IReadOnlyList<OfficeDocumentNestedResult> nested = input.NestedDocuments ?? Array.Empty<OfficeDocumentNestedResult>();
                for (int index = 0; index < nested.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    OfficeDocumentNestedResult entry = nested[index];
                    string childId = id + "/n" + index.ToString(CultureInfo.InvariantCulture);
                    string? childPath = ResolveNestedPath(path, input.Source?.Path, entry.Path);
                    OfficeDocumentReadResult child = entry.Document;
                    if (active.Contains(child)) throw new ArgumentException("Nested OCR requires an acyclic document tree.", nameof(document));
                    if (depth >= effective.MaxNestedDepth || report.DocumentCount >= effective.MaxDocuments) {
                        report.SkippedDocumentCount++;
                        var diagnostic = new OfficeDocumentDiagnostic {
                            Code = "ocr-nested-document-limit", Category = OfficeDocumentDiagnosticCategory.Limit,
                            Severity = OfficeDocumentDiagnosticSeverity.Warning, Source = execution.Id, IsRecoverable = true,
                            Message = "Nested OCR was skipped at the configured document or depth limit; source coverage remains incomplete.",
                            Location = new ReaderLocation { Path = childPath }
                        };
                        diagnostics.Add(diagnostic);
                        output.Diagnostics = output.Diagnostics.Concat(new[] { diagnostic }).ToArray();
                    } else {
                        child = await Visit(child, childId, childPath, entry.Path, depth + 1).ConfigureAwait(false);
                        AppendNestedOcrProjection(output, entry.Document, child, childId, childPath, entry.Path,
                            effective.EnrichmentOptions.AppendRecognizedTextToMarkdown, cancellationToken);
                    }
                    nestedResults.Add(new OfficeDocumentNestedResult { Path = entry.Path, Document = child });
                }
                output.NestedDocuments = nestedResults;
                return output;
            } finally {
                active.Remove(input);
            }
        }

        OfficeDocumentReadResult enriched = await Visit(document, "root", document.Source?.Path, document.Source?.Path ?? string.Empty, 0).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        enriched.Metadata = BuildExecutionMetadata(enriched.Metadata, report);
        return new OfficeDocumentOcrExecutionResult { Document = enriched, Recognitions = recognitions,
            Diagnostics = diagnostics, Report = report };
    }

    private static string? ResolveNestedPath(string? parent, string? originalParent, string child) {
        if (string.IsNullOrWhiteSpace(parent)) return string.IsNullOrWhiteSpace(child) ? null : child;
        if (string.IsNullOrWhiteSpace(child)) return parent;
        if (IsNestedPath(parent!, child) || System.IO.Path.IsPathRooted(child)) return child;
        if (originalParent != null && IsNestedPath(originalParent, child)) return parent + child.Substring(originalParent.Length);
        return parent + "!/" + child;
    }

    private static bool IsNestedPath(string parent, string child) =>
        child.StartsWith(parent + "!/", StringComparison.Ordinal) || child.StartsWith(parent + "::", StringComparison.Ordinal);

    private static void AddReport(OfficeDocumentOcrExecutionReport total, OfficeDocumentOcrExecutionReport own) {
        total.DocumentCount += own.DocumentCount;
        total.CandidateCount += own.CandidateCount;
        total.SelectedCandidateCount += own.SelectedCandidateCount;
        total.AttemptedCandidateCount += own.AttemptedCandidateCount;
        total.RecognizedCandidateCount += own.RecognizedCandidateCount;
        total.EmptyCandidateCount += own.EmptyCandidateCount;
        total.SkippedCandidateCount += own.SkippedCandidateCount;
        total.FailedCandidateCount += own.FailedCandidateCount;
        total.LineSpanCount += own.LineSpanCount;
        total.WordSpanCount += own.WordSpanCount;
        total.CharacterSpanCount += own.CharacterSpanCount;
        total.InputBytes += own.InputBytes;
        total.EffectiveDegreeOfParallelism = Math.Max(total.EffectiveDegreeOfParallelism, own.EffectiveDegreeOfParallelism);
    }
}
