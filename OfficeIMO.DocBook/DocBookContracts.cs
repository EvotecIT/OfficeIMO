using OfficeIMO.Core;
using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.DocBook;

/// <summary>Exact writer and bounded-validation profiles supported by OfficeIMO.DocBook.</summary>
public enum DocBookProfile {
    /// <summary>DocBook XML 4.5 DTD profile.</summary>
    DocBook45,
    /// <summary>DocBook 5.2 non-XInclude RELAX NG plus Schematron profile.</summary>
    DocBook52
}

/// <summary>Supported document roots.</summary>
public enum DocBookDocumentKind {
    /// <summary>Article root.</summary>
    Article,
    /// <summary>Book root.</summary>
    Book
}

/// <summary>Common structures with typed authoring support.</summary>
public enum DocBookNodeKind {
    /// <summary>Preserved extension outside the typed profile.</summary>
    Unknown,
    /// <summary>Metadata container.</summary>
    Info,
    /// <summary>Title.</summary>
    Title,
    /// <summary>Subtitle.</summary>
    Subtitle,
    /// <summary>Author metadata.</summary>
    Author,
    /// <summary>Section.</summary>
    Section,
    /// <summary>Paragraph.</summary>
    Paragraph,
    /// <summary>Unordered list.</summary>
    ItemizedList,
    /// <summary>Ordered list.</summary>
    OrderedList,
    /// <summary>Variable or definition list.</summary>
    VariableList,
    /// <summary>List item.</summary>
    ListItem,
    /// <summary>Table.</summary>
    Table,
    /// <summary>CALS table group.</summary>
    TableGroup,
    /// <summary>Table head.</summary>
    TableHead,
    /// <summary>Table body.</summary>
    TableBody,
    /// <summary>Table row.</summary>
    Row,
    /// <summary>Table cell entry.</summary>
    Entry,
    /// <summary>Program listing.</summary>
    ProgramListing,
    /// <summary>Screen or terminal content.</summary>
    Screen,
    /// <summary>Link.</summary>
    Link,
    /// <summary>Cross-reference.</summary>
    CrossReference,
    /// <summary>Note admonition.</summary>
    Note,
    /// <summary>Tip admonition.</summary>
    Tip,
    /// <summary>Important admonition.</summary>
    Important,
    /// <summary>Caution admonition.</summary>
    Caution,
    /// <summary>Warning admonition.</summary>
    Warning,
    /// <summary>Figure.</summary>
    Figure,
    /// <summary>Media object.</summary>
    MediaObject,
    /// <summary>Image object.</summary>
    ImageObject,
    /// <summary>Image data.</summary>
    ImageData,
    /// <summary>Caption.</summary>
    Caption,
    /// <summary>Index.</summary>
    Index,
    /// <summary>Index term.</summary>
    IndexTerm
}

/// <summary>Official identifiers associated with one exact schema profile.</summary>
public sealed class DocBookSchemaProfile {
    /// <summary>OfficeIMO profile selector.</summary>
    public DocBookProfile Profile { get; }
    /// <summary>DocBook namespace, empty for 4.5.</summary>
    public string NamespaceUri { get; }
    /// <summary>DTD public identifier when applicable.</summary>
    public string? DtdPublicId { get; }
    /// <summary>DTD system identifier when applicable.</summary>
    public string? DtdSystemId { get; }
    /// <summary>Normative RELAX NG schema URI when applicable.</summary>
    public string? RelaxNgUri { get; }
    /// <summary>Normative Schematron rules URI when applicable.</summary>
    public string? SchematronUri { get; }

    internal DocBookSchemaProfile(DocBookProfile profile, string ns, string? publicId, string? systemId, string? rng, string? schematron) {
        Profile = profile; NamespaceUri = ns; DtdPublicId = publicId; DtdSystemId = systemId; RelaxNgUri = rng; SchematronUri = schematron;
    }
}

/// <summary>Exact schema identifiers exposed by the product.</summary>
public static class DocBookSchemaProfiles {
    /// <summary>DocBook XML 4.5 DTD profile.</summary>
    public static DocBookSchemaProfile DocBook45 { get; } = new DocBookSchemaProfile(
        DocBookProfile.DocBook45, string.Empty, "-//OASIS//DTD DocBook XML V4.5//EN",
        "http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd", null, null);

    /// <summary>DocBook 5.2 non-XInclude RELAX NG plus Schematron profile.</summary>
    public static DocBookSchemaProfile DocBook52 { get; } = new DocBookSchemaProfile(
        DocBookProfile.DocBook52, "http://docbook.org/ns/docbook", null, null,
        "https://docs.oasis-open.org/docbook/docbook/v5.2/os/rng/docbook.rng",
        "https://docs.oasis-open.org/docbook/docbook/v5.2/os/sch/docbook.sch");

    /// <summary>Returns the exact schema identifiers for a profile.</summary>
    public static DocBookSchemaProfile Get(DocBookProfile profile) {
        switch (profile) {
            case DocBookProfile.DocBook45: return DocBook45;
            case DocBookProfile.DocBook52: return DocBook52;
            default: throw new ArgumentOutOfRangeException(nameof(profile));
        }
    }
}

/// <summary>Controls secure DocBook parsing limits.</summary>
public sealed class DocBookReadOptions {
    /// <summary>Maximum encoded input size. Defaults to 32 MiB.</summary>
    public long MaxInputBytes { get; set; } = 32L * 1024L * 1024L;
    /// <summary>Maximum XML characters.</summary>
    public long MaxCharacters { get; set; } = 32_000_000L;
    /// <summary>Maximum XML element depth.</summary>
    public int MaxDepth { get; set; } = 256;
    /// <summary>Maximum element count.</summary>
    public int MaxElements { get; set; } = 250_000;
    /// <summary>Maximum attribute count.</summary>
    public int MaxAttributes { get; set; } = 1_000_000;
    /// <summary>Maximum characters produced by bounded internal general-entity expansion. External resolution is disabled; external and parameter entity declarations in internal subsets are rejected.</summary>
    public long MaxCharactersFromEntities { get; set; } = 4096;

    internal void Validate() {
        if (MaxInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
        if (MaxCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaxCharacters));
        if (MaxDepth < 1) throw new ArgumentOutOfRangeException(nameof(MaxDepth));
        if (MaxElements < 1) throw new ArgumentOutOfRangeException(nameof(MaxElements));
        if (MaxAttributes < 1) throw new ArgumentOutOfRangeException(nameof(MaxAttributes));
        if (MaxCharactersFromEntities < 1) throw new ArgumentOutOfRangeException(nameof(MaxCharactersFromEntities));
    }
}

/// <summary>Controls DocBook serialization.</summary>
public sealed class DocBookWriteOptions {
    /// <summary>Pretty-prints changed XML.</summary>
    public bool Indent { get; set; } = true;
    /// <summary>Emits an unchanged parsed document byte-for-byte.</summary>
    public bool PreserveUnchangedSource { get; set; } = true;
}

/// <summary>Controls bounded common-structure validation diagnostics.</summary>
public sealed class DocBookValidationOptions {
    /// <summary>Maximum detailed diagnostics retained for each diagnostic code before one summary is emitted. Defaults to 100.</summary>
    public int MaxDetailedDiagnosticsPerCode { get; set; } = 100;

    internal void Validate() {
        if (MaxDetailedDiagnosticsPerCode < 1) throw new ArgumentOutOfRangeException(nameof(MaxDetailedDiagnosticsPerCode));
    }
}

/// <summary>Controls bounded projections from native DocBook into shared semantic channels.</summary>
public sealed class DocBookConversionOptions {
    /// <summary>Maximum columns materialized for one shared flat-table projection. Defaults to 1,024.</summary>
    public int MaxTableColumns { get; set; } = 1_024;
    /// <summary>Maximum header rows and body rows inspected for one shared flat-table projection. Defaults to 100,000 of each.</summary>
    public int MaxTableRows { get; set; } = 100_000;
    /// <summary>Maximum detailed diagnostics retained for each diagnostic code before one summary is emitted. Defaults to 100.</summary>
    public int MaxDetailedDiagnosticsPerCode { get; set; } = 100;

    internal void Validate() {
        if (MaxTableColumns < 1) throw new ArgumentOutOfRangeException(nameof(MaxTableColumns));
        if (MaxTableRows < 1) throw new ArgumentOutOfRangeException(nameof(MaxTableRows));
        if (MaxDetailedDiagnosticsPerCode < 1) throw new ArgumentOutOfRangeException(nameof(MaxDetailedDiagnosticsPerCode));
    }
}

/// <summary>Diagnostic severity.</summary>
public enum DocBookDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Info,
    /// <summary>Potential profile issue or conversion loss.</summary>
    Warning,
    /// <summary>Invalid supported-profile content.</summary>
    Error
}

/// <summary>Stable validation or conversion diagnostic.</summary>
public sealed class DocBookDiagnostic {
    /// <summary>Machine-readable code.</summary>
    public string Code { get; }
    /// <summary>Severity.</summary>
    public DocBookDiagnosticSeverity Severity { get; }
    /// <summary>Message.</summary>
    public string Message { get; }
    /// <summary>Best-effort element path.</summary>
    public string? Path { get; }
    /// <summary>Creates a diagnostic.</summary>
    public DocBookDiagnostic(string code, DocBookDiagnosticSeverity severity, string message, string? path = null) {
        Code = code ?? throw new ArgumentNullException(nameof(code)); Severity = severity;
        Message = message ?? throw new ArgumentNullException(nameof(message)); Path = path;
    }
}

/// <summary>Scope of validation actually performed.</summary>
public enum DocBookValidationScope {
    /// <summary>The bounded OfficeIMO common-structure profile, not a complete external DTD/RNG/Schematron run.</summary>
    OfficeIMOCommonStructure
}

/// <summary>Result of bounded DocBook validation.</summary>
public sealed class DocBookValidationResult {
    /// <summary>Exact schema/profile identifiers selected for the document.</summary>
    public DocBookSchemaProfile SchemaProfile { get; }
    /// <summary>Validation scope. This product does not claim full vocabulary-extension validation.</summary>
    public DocBookValidationScope Scope => DocBookValidationScope.OfficeIMOCommonStructure;
    /// <summary>Always false: callers wanting official schema validation must run the exposed DTD or RNG/Schematron artifacts.</summary>
    public bool IsOfficialSchemaValidated => false;
    /// <summary>True when the supported common structure has no errors.</summary>
    public bool IsValid { get; }
    /// <summary>Deterministic diagnostics.</summary>
    public IReadOnlyList<DocBookDiagnostic> Diagnostics { get; }

    internal DocBookValidationResult(DocBookSchemaProfile profile, IReadOnlyList<DocBookDiagnostic> diagnostics) {
        SchemaProfile = profile; Diagnostics = diagnostics;
        IsValid = !System.Linq.Enumerable.Any(diagnostics, d => d.Severity == DocBookDiagnosticSeverity.Error);
    }
}

/// <summary>Conversion result with deterministic loss reporting.</summary>
public sealed class DocBookConversionResult<T> : IOfficeConversionReport {
    /// <summary>Converted value.</summary>
    public T Value { get; }
    /// <summary>Conversion diagnostics.</summary>
    public IReadOnlyList<DocBookDiagnostic> Diagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss { get; }
    internal DocBookConversionResult(T value, IReadOnlyList<DocBookDiagnostic> diagnostics) {
        Value = value; Diagnostics = diagnostics;
        HasLoss = System.Linq.Enumerable.Any(diagnostics, d => d.Severity != DocBookDiagnosticSeverity.Info);
    }
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidDataException("The DocBook conversion reported loss. Inspect Diagnostics for details.");
    }
}
