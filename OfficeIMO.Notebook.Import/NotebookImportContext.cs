using System.Globalization;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Notebook.Import;

internal sealed class NotebookImportContext(NotebookImportOptions options, string format, CancellationToken token) {
    private const int MaxDiagnosticDetails = 4_096;
    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
    private readonly List<OfficeDiagnostic> _diagnostics = new();
    private readonly HashSet<string> _ids = new(StringComparer.Ordinal);
    private int _outputCount;

    internal NotebookDocument Document { get; } = NotebookDocument.Create();
    internal NotebookImportOptions Options => options;
    internal NotebookLoadOptions Limits => options.Limits;
    internal CancellationToken CancellationToken => token;

    internal NotebookCell AddCell(NotebookCellKind kind, string source, string? language = null, string? sourceId = null) {
        token.ThrowIfCancellationRequested();
        int index = Document.Cells.Count;
        if (index >= Limits.MaxCells) throw NotebookParser.Limit(nameof(Limits.MaxCells), Limits.MaxCells);
        if (NotebookParser.IsValidId(sourceId) && _ids.Add(sourceId!)) {
            return Document.AddImportedCell(kind, source, language, sourceId!);
        }

        // Resolve generated IDs after all source IDs have been seen so an unnamed cell
        // cannot take the ID of a later, explicitly named cell.
        NotebookCell cell = Document.AddImportedCell(kind, source, language, string.Empty);
        if (sourceId != null) {
            cell.SetMetadata("officeimo_source_id", JsonString(sourceId));
            Report("notebook-import-cell-id", "The source cell ID is retained in metadata because it could not be used as a Jupyter cell ID.",
                "cells[" + index.ToString(CultureInfo.InvariantCulture) + "]", OfficeConversionLossKind.None);
        }
        return cell;
    }

    internal void AddOutput(NotebookCell cell, NotebookOutput output) {
        token.ThrowIfCancellationRequested();
        if (cell.Outputs.Count >= Limits.MaxOutputsPerCell) {
            throw NotebookParser.Limit(nameof(Limits.MaxOutputsPerCell), Limits.MaxOutputsPerCell);
        }
        if (_outputCount >= Limits.MaxTotalOutputs) throw NotebookParser.Limit(nameof(Limits.MaxTotalOutputs), Limits.MaxTotalOutputs);
        cell.AddOutput(output);
        _outputCount++;
    }

    internal void Report(string code, string message, string? location = null,
        OfficeConversionLossKind lossKind = OfficeConversionLossKind.Omission,
        OfficeDiagnosticSeverity severity = OfficeDiagnosticSeverity.Warning) {
        if (_diagnostics.Count < MaxDiagnosticDetails) {
            _diagnostics.Add(new OfficeDiagnostic(code, severity, message, location, format, lossKind));
        } else if (_diagnostics.Count == MaxDiagnosticDetails) {
            _diagnostics.Add(new OfficeDiagnostic("notebook-import-diagnostic-limit", OfficeDiagnosticSeverity.Warning,
                "Additional notebook import diagnostics were suppressed after reaching the report limit.",
                format: format, lossKind: OfficeConversionLossKind.Unassessed));
        }
    }

    internal OfficeConversionResult<NotebookDocument> Complete(TimeSpan elapsed) {
        token.ThrowIfCancellationRequested();
        for (int index = 0; index < Document.Cells.Count; index++) {
            token.ThrowIfCancellationRequested();
            NotebookCell cell = Document.Cells[index];
            if (cell.Id.Length != 0) continue;
            string prefix = "cell-" + (index + 1).ToString(CultureInfo.InvariantCulture);
            string id = prefix;
            for (int suffix = 2; !_ids.Add(id); suffix++) {
                id = prefix + "-" + suffix.ToString(CultureInfo.InvariantCulture);
            }
            cell.SetImportedId(id);
        }
        var report = new OfficeConversionReport(format, "Jupyter Notebook", itemsWritten: Document.Cells.Count,
            elapsed: elapsed, diagnostics: _diagnostics);
        Document.ApplyImportReport(report);
        return new OfficeConversionResult<NotebookDocument>(Document, report);
    }

    internal static string DecodeUtf8(ReadOnlyMemory<byte> input) {
        try {
#if NET8_0_OR_GREATER
            return StrictUtf8.GetString(input.Span);
#else
            if (System.Runtime.InteropServices.MemoryMarshal.TryGetArray(input, out ArraySegment<byte> segment)) {
                return StrictUtf8.GetString(segment.Array!, segment.Offset, segment.Count);
            }
            return StrictUtf8.GetString(input.ToArray());
#endif
        } catch (DecoderFallbackException error) {
            throw new FormatException("The notebook input contains invalid UTF-8.", error);
        }
    }

    internal static JsonElement JsonString(string value) => NotebookJson.StringValue(value);
}
