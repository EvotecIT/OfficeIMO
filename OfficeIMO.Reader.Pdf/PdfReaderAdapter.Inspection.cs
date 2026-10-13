using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static PdfInspectionReadResult? LoadInspection(PdfDocument pdf, ReaderOptions readerOptions,
        ReaderPdfOptions pdfOptions, CancellationToken cancellationToken) => readerOptions.Inspection is null
        ? null : pdf.ReadForInspection(pdfOptions.ReadOptions, cancellationToken);

    private static IEnumerable<ReaderChunk> BuildInspectionWarningChunks(PdfInspectionReadResult? inspection,
        SourceMetadata source) {
        if (inspection is null) yield break;
        foreach (PdfPageReadOutcome outcome in inspection.PageOutcomes) {
            if (outcome.Succeeded) continue;
            string message = GetInspectionFailureMessage(outcome);
            yield return EnrichChunk(new ReaderChunk {
                Id = "pdf-inspection-failure-" + outcome.SelectionIndex.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Pdf,
                Location = InspectionLocation(source, outcome.PageNumber, outcome.SelectionIndex),
                Text = message,
                Warnings = new[] { message }
            }, source, computeHashes: false);
        }
    }

    private static OfficeDocumentReadResult AddInspectionEvidence(OfficeDocumentReadResult result,
        IReadOnlyList<PdfLogicalPage> pages, PdfInspectionReadResult? inspection, ReaderOptions options,
        SourceMetadata source, CancellationToken cancellationToken) {
        if (options.Inspection is null) return result;
        var evidence = new ReaderInspectionMetadataBuilder(options, "pdf", cancellationToken);
        evidence.Coverage("NativePageText", inspection is null ? "existing-projection" : inspection.CompletedAllSelectedPages ? "completed" : "partial",
            inspection is null ? "supplied logical pages; original source completion is not reassessed"
                : "selected source page occurrences; native text only; OCR is not performed");
        var diagnostics = result.Diagnostics.ToList();
        if (inspection is not null) {
            evidence.Add("pdf.source.structure", "SourcePageCount", inspection.SourceInfo.PageCount.ToString(CultureInfo.InvariantCulture), "count");
            foreach (PdfPageReadOutcome outcome in inspection.PageOutcomes) {
                cancellationToken.ThrowIfCancellationRequested();
                ReaderLocation location = InspectionLocation(source, outcome.PageNumber, outcome.SelectionIndex);
                var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["selectionIndex"] = outcome.SelectionIndex.ToString(CultureInfo.InvariantCulture)
                };
                if (outcome.MappingFailure is PdfUnsupportedTextMappingException mapping) {
                    attributes["fontResource"] = mapping.FontResource;
                    attributes["baseFont"] = mapping.BaseFont;
                    attributes["encoding"] = mapping.EncodingName;
                    attributes["hasDeclaredToUnicode"] = mapping.HasDeclaredToUnicode.ToString();
                }
                if (outcome.UnsupportedFilters.Count > 0) attributes["unsupportedFilters"] = string.Join(",", outcome.UnsupportedFilters);
                evidence.Add("pdf.page-outcome", "NativeText", outcome.Status.ToString(), "outcome", location, attributes: attributes);
                if (!outcome.Succeeded) {
                    diagnostics.Add(new OfficeDocumentDiagnostic {
                        Code = outcome.Status == PdfPageReadStatus.UnmappableText ? "pdf-page-unmappable-text" : "pdf-page-unsupported-content-filter",
                        Severity = OfficeDocumentDiagnosticSeverity.Error,
                        Message = GetInspectionFailureMessage(outcome),
                        Location = location,
                        Attributes = attributes
                    });
                    AddUnreadablePageStructure(evidence, inspection.SourceInfo.Pages[outcome.PageNumber - 1], location);
                }
            }
            foreach (PdfRepairDiagnostic repair in inspection.RepairReport.Diagnostics) {
                cancellationToken.ThrowIfCancellationRequested();
                string? objectId = repair.ObjectNumber?.ToString(CultureInfo.InvariantCulture);
                evidence.Add("pdf.repair", repair.Code, repair.Message, "repair", sourceObjectId: objectId,
                    attributes: new Dictionary<string, string>(StringComparer.Ordinal) { ["disposition"] = repair.Disposition.ToString() });
                diagnostics.Add(new OfficeDocumentDiagnostic {
                    Code = "pdf-repair-" + repair.Code,
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Message = repair.Message,
                    Attributes = new Dictionary<string, string>(StringComparer.Ordinal) { ["disposition"] = repair.Disposition.ToString() }
                });
            }
        }
        int selectedIndex = 0;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            if (inspection is not null) {
                while (!inspection.PageOutcomes[selectedIndex].Succeeded) selectedIndex++;
            }
            IReadOnlyList<PdfTextSpan> runs = page.Analysis.DecodedRuns;
            for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfTextSpan run = runs[runIndex];
                PdfTextDecodingEvidence decoding = run.DecodingEvidence;
                var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["source"] = decoding.Source.ToString(),
                    ["selectionIndex"] = (inspection is null ? pageIndex : selectedIndex).ToString(CultureInfo.InvariantCulture),
                    ["coordinateSpace"] = "pdf-user-space",
                    ["fontResource"] = run.FontResource,
                    ["usedFallback"] = decoding.UsedFallback.ToString(),
                    ["unmappedCodeBytes"] = decoding.UnmappedCodeBytes.ToString(CultureInfo.InvariantCulture),
                    ["hasDeclaredToUnicode"] = decoding.HasDeclaredToUnicode.ToString(),
                    ["usesDifferences"] = decoding.UsesDifferences.ToString(),
                    ["x"] = run.X.ToString("R", CultureInfo.InvariantCulture),
                    ["y"] = run.Y.ToString("R", CultureInfo.InvariantCulture)
                };
                if (decoding.EncodingName is not null) attributes["encoding"] = decoding.EncodingName;
                if (decoding.UnderlyingSource is PdfTextDecodingSource underlying) attributes["underlyingSource"] = underlying.ToString();
                evidence.Add("pdf.text-decoding", "NativeRun", run.Text, "decoded-text",
                    new ReaderLocation { Path = source.Path, Page = page.PageNumber, SourceBlockIndex = runIndex, SourceBlockKind = "text-run" },
                    run.ContentStreamObjectNumber?.ToString(CultureInfo.InvariantCulture), attributes);
                if (decoding.UsedFallback) {
                    diagnostics.Add(new OfficeDocumentDiagnostic {
                        Code = "pdf-text-decoding-fallback",
                        Severity = OfficeDocumentDiagnosticSeverity.Warning,
                        Message = "Native PDF text includes characters emitted through a fallback mapping. Inspect the located decoding evidence.",
                        Location = new ReaderLocation { Path = source.Path, Page = page.PageNumber, SourceBlockIndex = runIndex, SourceBlockKind = "text-run" },
                        Attributes = attributes
                    });
                }
            }
            selectedIndex++;
        }
        result.Diagnostics = diagnostics.AsReadOnly();
        return evidence.Complete(result);
    }

    private static void AddUnreadablePageStructure(ReaderInspectionMetadataBuilder evidence, PdfPageInfo page,
        ReaderLocation location) {
        evidence.Add("pdf.source.page", "Geometry", "retained", "source-structure", location,
            attributes: new Dictionary<string, string>(StringComparer.Ordinal) {
                ["width"] = page.Width.ToString("R", CultureInfo.InvariantCulture),
                ["height"] = page.Height.ToString("R", CultureInfo.InvariantCulture),
                ["rotation"] = page.RotationDegrees.ToString(CultureInfo.InvariantCulture)
            });
        foreach (PdfLinkAnnotation link in page.LinkAnnotations) {
            evidence.Add("pdf.source.link", "Target", link.Uri ?? link.DestinationName ?? link.NamedAction ?? link.RemoteFile,
                "inert-link", location, attributes: new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["destinationName"] = link.DestinationName ?? string.Empty,
                    ["destinationPage"] = link.DestinationPageNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                    ["namedAction"] = link.NamedAction ?? string.Empty,
                    ["remoteFile"] = link.RemoteFile ?? string.Empty,
                    ["remoteDestinationName"] = link.RemoteDestinationName ?? string.Empty,
                    ["remoteDestinationPage"] = link.RemoteDestinationPageNumber?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                    ["x1"] = link.X1.ToString("R", CultureInfo.InvariantCulture),
                    ["y1"] = link.Y1.ToString("R", CultureInfo.InvariantCulture),
                    ["x2"] = link.X2.ToString("R", CultureInfo.InvariantCulture),
                    ["y2"] = link.Y2.ToString("R", CultureInfo.InvariantCulture)
                });
        }
        foreach (PdfAnnotation annotation in page.Annotations) {
            evidence.Add("pdf.source.annotation", annotation.Subtype, annotation.Contents, "annotation", location,
                annotation.ObjectNumber?.ToString(CultureInfo.InvariantCulture),
                new Dictionary<string, string>(StringComparer.Ordinal) { ["actionType"] = annotation.ActionType ?? string.Empty });
        }
        foreach (PdfPageAction action in page.PageActions) {
            evidence.Add("pdf.source.page-action", action.TriggerName, action.ActionType, "inert-action", location,
                attributes: new Dictionary<string, string>(StringComparer.Ordinal) { ["path"] = action.ActionPath, ["uri"] = action.Uri ?? string.Empty });
        }
    }

    private static ReaderLocation InspectionLocation(SourceMetadata source, int pageNumber, int selectionIndex) => new ReaderLocation {
        Path = source.Path, Page = pageNumber, SourceBlockIndex = selectionIndex,
        SourceBlockKind = "page-outcome", BlockAnchor = "page-" + pageNumber.ToString(CultureInfo.InvariantCulture)
            + "-selection-" + selectionIndex.ToString("D4", CultureInfo.InvariantCulture)
    };

    private static string GetInspectionFailureMessage(PdfPageReadOutcome outcome) =>
        "PDF source page " + outcome.PageNumber.ToString(CultureInfo.InvariantCulture) + " was not read: " +
        (outcome.MappingFailure?.Message ?? "unsupported content-stream filters " + string.Join(", ", outcome.UnsupportedFilters)) + ".";
}
