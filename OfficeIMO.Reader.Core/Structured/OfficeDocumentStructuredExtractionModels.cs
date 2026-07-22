using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

/// <summary>Version identifiers for schema-friendly structured extraction output.</summary>
public static class OfficeDocumentStructuredExtractionSchema {
    /// <summary>Structured extraction schema identifier.</summary>
    public const string Id = "officeimo.document.structured-extraction";

    /// <summary>Current structured extraction schema version.</summary>
    public const int Version = 1;
}

/// <summary>Bounded options for deterministic structured extraction.</summary>
public sealed class OfficeDocumentStructuredExtractionOptions {
    /// <summary>Include document metadata as scalar records.</summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>Include form fields as scalar records and typed form entries.</summary>
    public bool IncludeForms { get; set; } = true;

    /// <summary>Extract rows from two-column and Path/Type/Value tables.</summary>
    public bool IncludeKeyValueRows { get; set; } = true;

    /// <summary>Extract Visio shape-data rows.</summary>
    public bool IncludeShapeData { get; set; } = true;

    /// <summary>Include deterministic chart summaries.</summary>
    public bool IncludeChartSummaries { get; set; } = true;

    /// <summary>Include visual, table, and chunk readiness summaries.</summary>
    public bool IncludeQualitySummaries { get; set; } = true;

    /// <summary>Build heading-and-following-content sections.</summary>
    public bool IncludeSections { get; set; } = true;

    /// <summary>Include tables with non-empty titles.</summary>
    public bool IncludeNamedTables { get; set; } = true;

    /// <summary>Copy source diagnostics into the extraction result.</summary>
    public bool IncludeSourceDiagnostics { get; set; } = true;

    /// <summary>Maximum scalar records. Default: 10,000.</summary>
    public int MaxRecords { get; set; } = 10_000;

    /// <summary>Maximum sections. Default: 1,000.</summary>
    public int MaxSections { get; set; } = 1_000;

    /// <summary>Maximum characters retained in one section. Default: 100,000.</summary>
    public int MaxSectionCharacters { get; set; } = 100_000;

    /// <summary>Maximum named tables. Default: 1,000.</summary>
    public int MaxTables { get; set; } = 1_000;

    /// <summary>Maximum typed forms. Default: 5,000.</summary>
    public int MaxForms { get; set; } = 5_000;

    /// <summary>Maximum copied source diagnostics. Default: 1,000.</summary>
    public int MaxDiagnostics { get; set; } = 1_000;

    /// <summary>Maximum characters parsed from one chart JSON payload. Default: 1,000,000.</summary>
    public int MaxChartContentCharacters { get; set; } = 1_000_000;

    internal OfficeDocumentStructuredExtractionOptions Clone() => new OfficeDocumentStructuredExtractionOptions {
        IncludeMetadata = IncludeMetadata,
        IncludeForms = IncludeForms,
        IncludeKeyValueRows = IncludeKeyValueRows,
        IncludeShapeData = IncludeShapeData,
        IncludeChartSummaries = IncludeChartSummaries,
        IncludeQualitySummaries = IncludeQualitySummaries,
        IncludeSections = IncludeSections,
        IncludeNamedTables = IncludeNamedTables,
        IncludeSourceDiagnostics = IncludeSourceDiagnostics,
        MaxRecords = MaxRecords,
        MaxSections = MaxSections,
        MaxSectionCharacters = MaxSectionCharacters,
        MaxTables = MaxTables,
        MaxForms = MaxForms,
        MaxDiagnostics = MaxDiagnostics,
        MaxChartContentCharacters = MaxChartContentCharacters
    };
}

/// <summary>One schema-friendly scalar record extracted from the shared document model.</summary>
public sealed class OfficeDocumentStructuredRecord {
    /// <summary>Deterministic record id within the extraction result.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Stable record category such as metadata, form, key-value, shape-data, chart-summary, or quality-summary.</summary>
    public string Category { get; set; } = string.Empty;

    /// <summary>Record name or key.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Scalar value when available.</summary>
    public string? Value { get; set; }

    /// <summary>Normalized or source-provided value type.</summary>
    public string? ValueType { get; set; }

    /// <summary>Source object identifier when available.</summary>
    public string? SourceObjectId { get; set; }

    /// <summary>Source location when available.</summary>
    public ReaderLocation? Location { get; set; }

    /// <summary>Additional deterministic scalar attributes.</summary>
    public IReadOnlyDictionary<string, string> Attributes { get; set; } = new SortedDictionary<string, string>(StringComparer.Ordinal);
}

/// <summary>Heading and the following logical content up to the next heading.</summary>
public sealed class OfficeDocumentStructuredSection {
    /// <summary>Deterministic section id in source order.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>Heading text, or null for preamble content.</summary>
    public string? Heading { get; set; }

    /// <summary>Heading level when available.</summary>
    public int? Level { get; set; }

    /// <summary>Following block text joined with newlines and bounded by extraction options.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Source block ids participating in this section.</summary>
    public IReadOnlyList<string> BlockIds { get; set; } = Array.Empty<string>();

    /// <summary>Heading or first-content location.</summary>
    public ReaderLocation? Location { get; set; }

    /// <summary>True when section text was truncated by <see cref="OfficeDocumentStructuredExtractionOptions.MaxSectionCharacters"/>.</summary>
    public bool Truncated { get; set; }
}

/// <summary>Bounded, schema-friendly extraction of a shared document result.</summary>
public sealed class OfficeDocumentStructuredExtractionResult {
    /// <summary>Structured extraction schema identifier.</summary>
    public string SchemaId { get; set; } = OfficeDocumentStructuredExtractionSchema.Id;

    /// <summary>Structured extraction schema version.</summary>
    public int SchemaVersion { get; set; } = OfficeDocumentStructuredExtractionSchema.Version;

    /// <summary>Source document metadata.</summary>
    public OfficeDocumentSource Source { get; set; } = new OfficeDocumentSource();

    /// <summary>Scalar structured records in deterministic category/source order.</summary>
    public IReadOnlyList<OfficeDocumentStructuredRecord> Records { get; set; } = Array.Empty<OfficeDocumentStructuredRecord>();

    /// <summary>Heading-based sections in source order.</summary>
    public IReadOnlyList<OfficeDocumentStructuredSection> Sections { get; set; } = Array.Empty<OfficeDocumentStructuredSection>();

    /// <summary>Named tables in source order.</summary>
    public IReadOnlyList<ReaderTable> Tables { get; set; } = Array.Empty<ReaderTable>();

    /// <summary>Typed form fields in source order.</summary>
    public IReadOnlyList<OfficeDocumentFormField> Forms { get; set; } = Array.Empty<OfficeDocumentFormField>();

    /// <summary>Source and extraction-limit diagnostics.</summary>
    public IReadOnlyList<OfficeDocumentDiagnostic> Diagnostics { get; set; } = Array.Empty<OfficeDocumentDiagnostic>();
}
