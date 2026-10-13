using System.Globalization;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Projects native owner evidence into existing located metadata, with operation-wide accounting.</summary>
internal sealed class ReaderInspectionMetadataBuilder {
    private readonly List<OfficeDocumentMetadataEntry> _records = new();
    private readonly List<OfficeDocumentDiagnostic> _diagnostics = new();
    private readonly ReaderOperationBudget _budget;
    private readonly CancellationToken _cancellationToken;
    private readonly string _prefix;

    internal ReaderInspectionMetadataBuilder(ReaderOptions options, string prefix, CancellationToken cancellationToken) {
        Options = options.Inspection?.CloneValidated() ?? throw new ArgumentException("Inspection options are required.", nameof(options));
        _budget = ReaderReadScope.Current?.Budget ?? new ReaderOperationBudget(new ReaderResourceLimits(), Options);
        _budget.IncludeInspectionLimits(Options);
        _prefix = prefix;
        _cancellationToken = cancellationToken;
        Add(OfficeDocumentInspectionCategories.Coverage, "configuration", "requested", "inspection-status",
            attributes: Attributes(("comments", Options.IncludeComments ? "true" : "false"),
                ("revisions", Options.IncludeRevisions ? "true" : "false"), ("formulas", Options.IncludeFormulas ? "true" : "false"),
                ("visibility", Options.IncludeVisibility ? "true" : "false"), ("packageContent", Options.IncludePackageContent ? "true" : "false"),
                ("contentSafety", Options.IncludeContentSafety ? "true" : "false"),
                ("maxRecords", Options.MaxRecords.ToString(CultureInfo.InvariantCulture)),
                ("maxCharacters", Options.MaxCharacters.ToString(CultureInfo.InvariantCulture))));
    }

    internal ReaderInspectionOptions Options { get; }

    internal void Add(string category, string name, string? value, string valueType,
        ReaderLocation? location = null, string? sourceObjectId = null,
        IReadOnlyDictionary<string, string>? attributes = null) {
        _cancellationToken.ThrowIfCancellationRequested();
        var record = new OfficeDocumentMetadataEntry {
            Id = _prefix + "-inspection-" + _records.Count.ToString("D6", CultureInfo.InvariantCulture),
            Category = category, Name = name, Value = value, ValueType = valueType,
            SourceObjectId = sourceObjectId, Location = location,
            Attributes = attributes ?? new Dictionary<string, string>(StringComparer.Ordinal)
        };
        _budget.AddInspectionRecord(record);
        _records.Add(record);
    }

    internal void Coverage(string name, string status, string scope, ReaderLocation? location = null) =>
        Add(OfficeDocumentInspectionCategories.Coverage, name, status, "inspection-status", location,
            attributes: new Dictionary<string, string>(StringComparer.Ordinal) { ["scope"] = scope });

    internal void Diagnostic(string code, string message, ReaderLocation? location = null) {
        Add(OfficeDocumentInspectionCategories.Coverage, code, message, "diagnostic", location);
        _diagnostics.Add(new OfficeDocumentDiagnostic {
            Code = code, Message = message, Category = OfficeDocumentDiagnosticCategory.Content,
            Severity = OfficeDocumentDiagnosticSeverity.Warning, Source = "officeimo.reader.inspection",
            IsRecoverable = true, Location = location
        });
    }

    internal static IReadOnlyDictionary<string, string> Attributes(params (string Name, string? Value)[] pairs) {
        var result = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in pairs) if (pair.Value != null) result.Add(pair.Name, pair.Value);
        return result;
    }

    internal OfficeDocumentReadResult Complete(OfficeDocumentReadResult result) {
        _cancellationToken.ThrowIfCancellationRequested();
        result.Metadata = result.Metadata.Concat(_records).ToArray();
        result.Diagnostics = result.Diagnostics.Concat(_diagnostics).ToArray();
        result.CapabilitiesUsed = result.CapabilitiesUsed.Concat(new[] { "officeimo.reader.inspection-metadata.v1" }).Distinct(StringComparer.Ordinal).ToArray();
        return result;
    }
}
