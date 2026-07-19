using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Deterministically extracts bounded schema-friendly records from a shared document result.</summary>
public static partial class OfficeDocumentStructuredExtractor {
    /// <summary>Extracts sections, scalar records, named tables, forms, and diagnostics.</summary>
    public static OfficeDocumentStructuredExtractionResult Extract(
        OfficeDocumentReadResult document,
        OfficeDocumentStructuredExtractionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        OfficeDocumentStructuredExtractionOptions effective = Normalize(options);
        var state = new ExtractionState(effective, cancellationToken);
        state.CopySourceDiagnostics(document.Diagnostics);
        IReadOnlyList<OfficeDocumentFormField> forms = effective.IncludeForms
            ? SelectForms(document, state)
            : Array.Empty<OfficeDocumentFormField>();

        if (effective.IncludeMetadata) AddMetadata(document, state);
        if (effective.IncludeForms) AddForms(forms, state);
        if (effective.IncludeKeyValueRows) AddKeyValueRows(document, state);
        if (effective.IncludeShapeData) AddShapeData(document, state);
        if (effective.IncludeChartSummaries) AddChartSummaries(document, state);
        if (effective.IncludeQualitySummaries) AddQualitySummaries(document, state);

        IReadOnlyList<OfficeDocumentStructuredSection> sections = effective.IncludeSections
            ? BuildSections(document, state)
            : Array.Empty<OfficeDocumentStructuredSection>();
        IReadOnlyList<ReaderTable> tables = effective.IncludeNamedTables
            ? SelectNamedTables(document, state)
            : Array.Empty<ReaderTable>();
        return new OfficeDocumentStructuredExtractionResult {
            Source = document.Source ?? new OfficeDocumentSource(),
            Records = state.Records.Count == 0 ? Array.Empty<OfficeDocumentStructuredRecord>() : state.Records.ToArray(),
            Sections = sections,
            Tables = tables,
            Forms = forms,
            Diagnostics = state.Diagnostics.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : state.Diagnostics.ToArray()
        };
    }

    private static OfficeDocumentStructuredExtractionOptions Normalize(OfficeDocumentStructuredExtractionOptions? options) {
        OfficeDocumentStructuredExtractionOptions effective = (options ?? new OfficeDocumentStructuredExtractionOptions()).Clone();
        ValidatePositive(effective.MaxRecords, nameof(effective.MaxRecords));
        ValidatePositive(effective.MaxSections, nameof(effective.MaxSections));
        ValidatePositive(effective.MaxSectionCharacters, nameof(effective.MaxSectionCharacters));
        ValidatePositive(effective.MaxTables, nameof(effective.MaxTables));
        ValidatePositive(effective.MaxForms, nameof(effective.MaxForms));
        ValidatePositive(effective.MaxDiagnostics, nameof(effective.MaxDiagnostics));
        return effective;
    }

    private static void ValidatePositive(int value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name, value, "Structured extraction limits must be positive.");
    }

    private static IReadOnlyList<ReaderTable> SelectNamedTables(
        OfficeDocumentReadResult document,
        ExtractionState state) {
        var selected = new List<ReaderTable>();
        foreach (ReaderTable table in OfficeDocumentModelTraversal.Tables(document)) {
            state.CancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(table.Title)) continue;
            if (selected.Count >= state.Options.MaxTables) {
                state.AddLimitDiagnostic("structured-table-limit", state.Options.MaxTables, "named tables");
                break;
            }
            selected.Add(table);
        }
        return selected.Count == 0 ? Array.Empty<ReaderTable>() : selected.ToArray();
    }

    private static IReadOnlyList<OfficeDocumentFormField> SelectForms(
        OfficeDocumentReadResult document,
        ExtractionState state) {
        var selected = new List<OfficeDocumentFormField>();
        foreach (OfficeDocumentFormField form in OfficeDocumentModelTraversal.Forms(document)) {
            state.CancellationToken.ThrowIfCancellationRequested();
            if (selected.Count >= state.Options.MaxForms) {
                state.AddLimitDiagnostic("structured-form-limit", state.Options.MaxForms, "forms");
                break;
            }
            selected.Add(form);
        }
        return selected.Count == 0 ? Array.Empty<OfficeDocumentFormField>() : selected.ToArray();
    }

    internal sealed class ExtractionState {
        private readonly HashSet<string> _limitCodes = new HashSet<string>(StringComparer.Ordinal);

        internal ExtractionState(
            OfficeDocumentStructuredExtractionOptions options,
            CancellationToken cancellationToken) {
            Options = options;
            CancellationToken = cancellationToken;
        }

        internal OfficeDocumentStructuredExtractionOptions Options { get; }
        internal CancellationToken CancellationToken { get; }
        internal List<OfficeDocumentStructuredRecord> Records { get; } = new List<OfficeDocumentStructuredRecord>();
        internal List<OfficeDocumentDiagnostic> Diagnostics { get; } = new List<OfficeDocumentDiagnostic>();

        internal bool TryAdd(OfficeDocumentStructuredRecord record) {
            CancellationToken.ThrowIfCancellationRequested();
            if (Records.Count >= Options.MaxRecords) {
                AddLimitDiagnostic("structured-record-limit", Options.MaxRecords, "records");
                return false;
            }
            record.Attributes = SortAttributes(record.Attributes);
            Records.Add(record);
            return true;
        }

        internal void CopySourceDiagnostics(IReadOnlyList<OfficeDocumentDiagnostic>? source) {
            if (!Options.IncludeSourceDiagnostics || source == null || source.Count == 0) return;
            int count = Math.Min(source.Count, Options.MaxDiagnostics);
            for (int index = 0; index < count; index++) {
                CancellationToken.ThrowIfCancellationRequested();
                Diagnostics.Add(source[index]);
            }
            if (source.Count > count) AddLimitDiagnostic("structured-diagnostic-limit", Options.MaxDiagnostics, "diagnostics");
        }

        internal void AddLimitDiagnostic(string code, int limit, string noun) {
            if (!_limitCodes.Add(code)) return;
            Diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Limit,
                Code = code,
                Message = $"Structured extraction reached the configured {noun} limit ({limit.ToString(CultureInfo.InvariantCulture)}).",
                Source = "officeimo.reader.structured-extraction",
                IsRecoverable = true,
                Attributes = new SortedDictionary<string, string>(StringComparer.Ordinal) {
                    ["limit"] = limit.ToString(CultureInfo.InvariantCulture),
                    ["target"] = noun
                }
            });
        }

        private static IReadOnlyDictionary<string, string> SortAttributes(IReadOnlyDictionary<string, string>? attributes) {
            var sorted = new SortedDictionary<string, string>(StringComparer.Ordinal);
            if (attributes == null) return sorted;
            foreach (KeyValuePair<string, string> attribute in attributes) sorted[attribute.Key] = attribute.Value;
            return sorted;
        }
    }
}
