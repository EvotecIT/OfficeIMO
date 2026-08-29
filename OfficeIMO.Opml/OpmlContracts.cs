using OfficeIMO.Core;
using OfficeIMO.Core.Internal;
using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Opml;

/// <summary>OPML revisions supported for creation and canonical writing.</summary>
public enum OpmlVersion {
    /// <summary>OPML 1.0. A declared 1.1 document is read using this profile.</summary>
    Opml10,
    /// <summary>OPML 2.0.</summary>
    Opml20
}

/// <summary>Controls parsing resource limits.</summary>
public sealed class OpmlReadOptions {
    /// <summary>Maximum encoded input size. Defaults to 16 MiB.</summary>
    public long MaxInputBytes { get; set; } = 16L * 1024L * 1024L;
    /// <summary>Maximum XML character count. Defaults to 16 million.</summary>
    public long MaxCharacters { get; set; } = 16_000_000L;
    /// <summary>Maximum XML element depth.</summary>
    public int MaxDepth { get; set; } = 128;
    /// <summary>Maximum total number of XML elements, including extension elements.</summary>
    public int MaxElements { get; set; } = 250_000;
    /// <summary>Maximum number of outline elements, enforced while parsing before XML materialization.</summary>
    public int MaxOutlines { get; set; } = 100_000;
    /// <summary>Maximum total number of XML attributes.</summary>
    public int MaxAttributes { get; set; } = 500_000;

    internal void Validate() {
        if (MaxInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
        if (MaxCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaxCharacters));
        if (MaxDepth < 1) throw new ArgumentOutOfRangeException(nameof(MaxDepth));
        if (MaxElements < 1) throw new ArgumentOutOfRangeException(nameof(MaxElements));
        if (MaxOutlines < 1) throw new ArgumentOutOfRangeException(nameof(MaxOutlines));
        if (MaxAttributes < 1) throw new ArgumentOutOfRangeException(nameof(MaxAttributes));
    }
}

/// <summary>Controls OPML serialization.</summary>
public sealed class OpmlWriteOptions {
    /// <summary>When true, pretty-prints XML after a document has changed.</summary>
    public bool Indent { get; set; } = true;
    /// <summary>When true, an unchanged parsed document is emitted byte-for-byte.</summary>
    public bool PreserveUnchangedSource { get; set; } = true;
}

/// <summary>Controls bounded OPML validation diagnostics.</summary>
public sealed class OpmlValidationOptions {
    /// <summary>Maximum detailed diagnostics retained for each diagnostic code before one summary is emitted. Defaults to 100.</summary>
    public int MaxDetailedDiagnosticsPerCode { get; set; } = 100;

    internal void Validate() {
        if (MaxDetailedDiagnosticsPerCode < 1) throw new ArgumentOutOfRangeException(nameof(MaxDetailedDiagnosticsPerCode));
    }
}

/// <summary>Controls bounded OPML shared-model conversion diagnostics.</summary>
public sealed class OpmlConversionOptions {
    /// <summary>Maximum shared-model structure depth accepted by reverse conversion. Defaults to 128 and cannot exceed 256.</summary>
    public int MaxStructureDepth { get; set; } = 128;
    /// <summary>Maximum shared-model structure nodes accepted by reverse conversion. Defaults to 100,000.</summary>
    public int MaxStructureNodes { get; set; } = 100_000;
    /// <summary>Maximum detailed diagnostics retained for each diagnostic code before one summary is emitted. Defaults to 100.</summary>
    public int MaxDetailedDiagnosticsPerCode { get; set; } = 100;

    internal void Validate() {
        if (MaxStructureDepth < 1 || MaxStructureDepth > OfficeDocumentModelStructureTraversal.MaximumSupportedDepth)
            throw new ArgumentOutOfRangeException(nameof(MaxStructureDepth));
        if (MaxStructureNodes < 1) throw new ArgumentOutOfRangeException(nameof(MaxStructureNodes));
        if (MaxDetailedDiagnosticsPerCode < 1) throw new ArgumentOutOfRangeException(nameof(MaxDetailedDiagnosticsPerCode));
    }
}

/// <summary>Severity of an OPML validation or conversion diagnostic.</summary>
public enum OpmlDiagnosticSeverity {
    /// <summary>Informational diagnostic.</summary>
    Info,
    /// <summary>Potential compatibility or conversion loss.</summary>
    Warning,
    /// <summary>Invalid supported-profile content.</summary>
    Error
}

/// <summary>A stable OPML validation or conversion diagnostic.</summary>
public sealed class OpmlDiagnostic {
    /// <summary>Machine-readable code.</summary>
    public string Code { get; }
    /// <summary>Severity.</summary>
    public OpmlDiagnosticSeverity Severity { get; }
    /// <summary>Human-readable message.</summary>
    public string Message { get; }
    /// <summary>Best-effort element path.</summary>
    public string? Path { get; }

    /// <summary>Creates a diagnostic.</summary>
    public OpmlDiagnostic(string code, OpmlDiagnosticSeverity severity, string message, string? path = null) {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Severity = severity;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        Path = path;
    }
}

/// <summary>Result of OPML profile validation.</summary>
public sealed class OpmlValidationResult {
    /// <summary>Effective profile used for validation.</summary>
    public OpmlVersion Profile { get; }
    /// <summary>Diagnostics in deterministic document order.</summary>
    public IReadOnlyList<OpmlDiagnostic> Diagnostics { get; }
    /// <summary>True when no error diagnostic was emitted.</summary>
    public bool IsValid { get; }

    internal OpmlValidationResult(OpmlVersion profile, IReadOnlyList<OpmlDiagnostic> diagnostics) {
        Profile = profile;
        Diagnostics = diagnostics;
        IsValid = !System.Linq.Enumerable.Any(diagnostics, d => d.Severity == OpmlDiagnosticSeverity.Error);
    }
}

/// <summary>Conversion result with deterministic loss reporting.</summary>
public sealed class OpmlConversionResult<T> : IOfficeConversionReport {
    /// <summary>Converted value.</summary>
    public T Value { get; }
    /// <summary>Conversion diagnostics.</summary>
    public IReadOnlyList<OpmlDiagnostic> Diagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss { get; }

    internal OpmlConversionResult(T value, IReadOnlyList<OpmlDiagnostic> diagnostics) {
        Value = value;
        Diagnostics = diagnostics;
        HasLoss = System.Linq.Enumerable.Any(diagnostics, d => d.Severity != OpmlDiagnosticSeverity.Info);
    }

    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidDataException("The OPML conversion reported loss. Inspect Diagnostics for details.");
    }
}
