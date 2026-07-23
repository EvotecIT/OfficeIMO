using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentStructuredExtractor {
    private static void AddChartSummaries(OfficeDocumentReadResult document, ExtractionState state) {
        int chartIndex = 0;
        foreach (ReaderVisual visual in EnumerateVisuals(document)) {
            state.CancellationToken.ThrowIfCancellationRequested();
            if (!IsChart(visual)) continue;
            if (!state.CanAddRecord()) return;
            var attributes = new SortedDictionary<string, string>(StringComparer.Ordinal);
            AddIfPresent(attributes, "language", visual.Language);
            AddIfPresent(attributes, "mimeType", visual.MimeType);
            AddIfPresent(attributes, "payloadHash", visual.PayloadHash);
            string? chartType = TryReadChartSummary(visual.Content, attributes, state);
            if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                Id = "chart-summary-" + chartIndex.ToString("D4", CultureInfo.InvariantCulture),
                Category = "chart-summary",
                Name = string.IsNullOrWhiteSpace(visual.SourceName) ? "Chart" : visual.SourceName!,
                Value = chartType ?? visual.Kind,
                ValueType = "chart",
                SourceObjectId = visual.PayloadHash,
                Location = visual.Location,
                Attributes = attributes
            })) return;
            chartIndex++;
        }
    }

    private static void AddQualitySummaries(OfficeDocumentReadResult document, ExtractionState state) {
        AddTableQualitySummaries(document, state);
        AddVisualQualitySummaries(document, state);
        AddChunkReadinessSummaries(document, state);
        AddDiagnosticReadinessSummaries(document, state);
    }

    private static void AddTableQualitySummaries(OfficeDocumentReadResult document, ExtractionState state) {
        int index = 0;
        foreach (ReaderTable table in OfficeDocumentModelTraversal.Tables(document)) {
            state.CancellationToken.ThrowIfCancellationRequested();
            ReaderTableDiagnostics? diagnostics = table.Diagnostics;
            if (diagnostics == null) {
                index++;
                continue;
            }
            var attributes = new SortedDictionary<string, string>(StringComparer.Ordinal) {
                ["cellCompleteness"] = Format(diagnostics.CellCompleteness),
                ["columnGeometryConfidence"] = Format(diagnostics.ColumnGeometryConfidence),
                ["filledCellCount"] = diagnostics.FilledCellCount.ToString(CultureInfo.InvariantCulture),
                ["hasGeometry"] = diagnostics.HasGeometry ? "true" : "false",
                ["missingCellCount"] = diagnostics.MissingCellCount.ToString(CultureInfo.InvariantCulture),
                ["schemaConfidence"] = Format(diagnostics.SchemaConfidence),
                ["sourceRowCount"] = diagnostics.SourceRowCount.ToString(CultureInfo.InvariantCulture)
            };
            if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                Id = "table-quality-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "quality-summary",
                Name = string.IsNullOrWhiteSpace(table.Title) ? "Table " + (index + 1).ToString(CultureInfo.InvariantCulture) : table.Title!,
                Value = Format(diagnostics.Confidence),
                ValueType = "confidence",
                Location = table.Location,
                Attributes = attributes
            })) return;
            index++;
        }
    }

    private static void AddVisualQualitySummaries(OfficeDocumentReadResult document, ExtractionState state) {
        int index = 0;
        foreach (ReaderVisual visual in EnumerateVisuals(document)) {
            state.CancellationToken.ThrowIfCancellationRequested();
            var attributes = new SortedDictionary<string, string>(StringComparer.Ordinal) {
                ["hasGeometry"] = visual.HasGeometry ? "true" : "false",
                ["isAxisAligned"] = visual.IsAxisAligned.HasValue
                    ? (visual.IsAxisAligned.Value ? "true" : "false")
                    : "unknown",
                ["placementCount"] = visual.PlacementCount.ToString(CultureInfo.InvariantCulture)
            };
            AddNullableNumber(attributes, "height", visual.Height);
            AddNullableNumber(attributes, "width", visual.Width);
            AddIfPresent(attributes, "language", visual.Language);
            AddIfPresent(attributes, "mimeType", visual.MimeType);
            if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                Id = "visual-quality-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "visual-summary",
                Name = string.IsNullOrWhiteSpace(visual.SourceName) ? visual.Kind ?? "visual" : visual.SourceName!,
                Value = visual.Kind,
                ValueType = "visual-kind",
                SourceObjectId = visual.PayloadHash,
                Location = visual.Location,
                Attributes = attributes
            })) return;
            index++;
        }
    }

    private static void AddChunkReadinessSummaries(OfficeDocumentReadResult document, ExtractionState state) {
        IReadOnlyList<ReaderChunk> chunks = document.Chunks ?? Array.Empty<ReaderChunk>();
        for (int index = 0; index < chunks.Count; index++) {
            state.CancellationToken.ThrowIfCancellationRequested();
            ReaderChunk chunk = chunks[index];
            ReaderChunkDiagnostics? diagnostics = chunk.Diagnostics;
            if (diagnostics == null) continue;
            var attributes = new SortedDictionary<string, string>(StringComparer.Ordinal) {
                ["attachmentCount"] = diagnostics.AttachmentCount.ToString(CultureInfo.InvariantCulture),
                ["formFieldCount"] = diagnostics.FormFieldCount.ToString(CultureInfo.InvariantCulture),
                ["hasActiveContent"] = diagnostics.HasActiveContent ? "true" : "false",
                ["hasEncryption"] = diagnostics.HasEncryption ? "true" : "false",
                ["hasSecurityState"] = diagnostics.HasSecurityState ? "true" : "false",
                ["hasSignatures"] = diagnostics.HasSignatures ? "true" : "false",
                ["imageCount"] = diagnostics.ImageCount.ToString(CultureInfo.InvariantCulture),
                ["linkCount"] = diagnostics.LinkCount.ToString(CultureInfo.InvariantCulture),
                ["potentiallyUnsafeActionCount"] = diagnostics.PotentiallyUnsafeActionCount.ToString(CultureInfo.InvariantCulture),
                ["tableCount"] = diagnostics.TableCount.ToString(CultureInfo.InvariantCulture)
            };
            if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                Id = "readiness-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "readiness-summary",
                Name = string.IsNullOrWhiteSpace(diagnostics.SourceKind) ? chunk.Id : diagnostics.SourceKind,
                Value = diagnostics.HasActiveContent || diagnostics.HasSecurityState ? "review" : "ready",
                ValueType = "readiness",
                SourceObjectId = chunk.Id,
                Location = chunk.Location,
                Attributes = attributes
            })) return;
        }
    }

    private static void AddDiagnosticReadinessSummaries(OfficeDocumentReadResult document, ExtractionState state) {
        IReadOnlyList<OfficeDocumentDiagnostic> diagnostics = document.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>();
        for (int index = 0; index < diagnostics.Count; index++) {
            state.CancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentDiagnostic diagnostic = diagnostics[index];
            if (!IsReadinessDiagnostic(diagnostic)) continue;
            var attributes = CopyAttributes(diagnostic.Attributes);
            AddIfPresent(attributes, "source", diagnostic.Source);
            attributes["severity"] = diagnostic.Severity.ToString();
            if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                Id = "diagnostic-readiness-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "readiness-summary",
                Name = diagnostic.Code,
                Value = diagnostic.Message,
                ValueType = diagnostic.Category.ToString(),
                Location = diagnostic.Location,
                Attributes = attributes
            })) return;
        }
    }

    private static IEnumerable<ReaderVisual> EnumerateVisuals(OfficeDocumentReadResult document) {
        foreach (ReaderVisual visual in OfficeDocumentModelTraversal.Visuals(document)) yield return visual;
    }

    private static bool IsChart(ReaderVisual visual) =>
        string.Equals(visual.Kind?.Trim(), "chart", StringComparison.OrdinalIgnoreCase) ||
        (!string.IsNullOrWhiteSpace(visual.Language) && visual.Language!.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0);

    private static string? TryReadChartSummary(string? content, IDictionary<string, string> attributes, ExtractionState state) {
        if (string.IsNullOrWhiteSpace(content)) return null;
        attributes["contentLength"] = content!.Length.ToString(CultureInfo.InvariantCulture);
        if (content.Length > state.Options.MaxChartContentCharacters) {
            attributes["contentSkipped"] = "true";
            state.AddLimitDiagnostic("structured-chart-content-limit", state.Options.MaxChartContentCharacters, "chart content characters");
            return null;
        }
        try {
            using JsonDocument json = JsonDocument.Parse(content, new JsonDocumentOptions { MaxDepth = 64 });
            JsonElement root = json.RootElement;
            string? type = root.TryGetProperty("type", out JsonElement typeElement) && typeElement.ValueKind == JsonValueKind.String
                ? typeElement.GetString()
                : null;
            if (root.TryGetProperty("data", out JsonElement data) && data.ValueKind == JsonValueKind.Object) {
                if (data.TryGetProperty("labels", out JsonElement labels) && labels.ValueKind == JsonValueKind.Array) {
                    attributes["labelCount"] = labels.GetArrayLength().ToString(CultureInfo.InvariantCulture);
                }
                if (data.TryGetProperty("datasets", out JsonElement datasets) && datasets.ValueKind == JsonValueKind.Array) {
                    attributes["datasetCount"] = datasets.GetArrayLength().ToString(CultureInfo.InvariantCulture);
                    long pointCount = 0;
                    foreach (JsonElement dataset in datasets.EnumerateArray()) {
                        state.CancellationToken.ThrowIfCancellationRequested();
                        if (dataset.ValueKind == JsonValueKind.Object &&
                            dataset.TryGetProperty("data", out JsonElement points) &&
                            points.ValueKind == JsonValueKind.Array) {
                            pointCount = checked(pointCount + points.GetArrayLength());
                        }
                    }
                    attributes["pointCount"] = pointCount.ToString(CultureInfo.InvariantCulture);
                }
            }
            return type;
        } catch (JsonException) {
            attributes["jsonValid"] = "false";
            return null;
        }
    }

    private static bool IsReadinessDiagnostic(OfficeDocumentDiagnostic diagnostic) {
        if (diagnostic.Category == OfficeDocumentDiagnosticCategory.Security ||
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Ocr ||
            diagnostic.Category == OfficeDocumentDiagnosticCategory.Limit) return true;
        string code = diagnostic.Code ?? string.Empty;
        return code.IndexOf("compliance", StringComparison.OrdinalIgnoreCase) >= 0 ||
               code.IndexOf("preflight", StringComparison.OrdinalIgnoreCase) >= 0 ||
               code.IndexOf("quality", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string Format(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    private static void AddNullableNumber(IDictionary<string, string> attributes, string name, double? value) {
        if (value.HasValue) attributes[name] = Format(value.Value);
    }

    private static void AddNullableNumber(IDictionary<string, string> attributes, string name, int? value) {
        if (value.HasValue) attributes[name] = value.Value.ToString(CultureInfo.InvariantCulture);
    }
}
