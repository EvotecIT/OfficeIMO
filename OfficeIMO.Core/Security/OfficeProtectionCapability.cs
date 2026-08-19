using System;

namespace OfficeIMO.Security;

/// <summary>Classifies a protected-content mechanism without conflating access hints with encryption.</summary>
public enum OfficeProtectionKind {
    /// <summary>Password-derived cryptographic content encryption.</summary>
    PasswordEncryption,
    /// <summary>Recipient-certificate cryptographic content encryption.</summary>
    RecipientEncryption,
    /// <summary>Certificate-backed or key-backed digital signatures.</summary>
    DigitalSignature,
    /// <summary>Standards-defined reversible font obfuscation.</summary>
    FontObfuscation,
    /// <summary>Editing restrictions that do not provide confidentiality.</summary>
    EditingRestriction,
    /// <summary>Checksum-based access deterrence that does not provide confidentiality.</summary>
    AccessDeterrence
}

/// <summary>Coverage state for one protected-content operation.</summary>
public enum OfficeProtectionCoverageState {
    /// <summary>The operation is implemented and contract-tested.</summary>
    Supported,
    /// <summary>The mechanism is identified but not opened, created, or validated.</summary>
    Detected,
    /// <summary>Unchanged protected bytes are retained.</summary>
    Preserved,
    /// <summary>The operation deliberately fails to prevent protection loss or stale-content exposure.</summary>
    Blocked,
    /// <summary>The operation is not implemented.</summary>
    NotSupported,
    /// <summary>The operation does not apply to this mechanism.</summary>
    NotApplicable
}

/// <summary>One machine-readable protected-content capability row.</summary>
public sealed class OfficeProtectionCapability {
    /// <summary>Creates one protected-content capability row.</summary>
    public OfficeProtectionCapability(string id, string formatId, string packageId, OfficeProtectionKind kind,
        OfficeProtectionCoverageState inspect, OfficeProtectionCoverageState open,
        OfficeProtectionCoverageState create, OfficeProtectionCoverageState validate,
        OfficeProtectionCoverageState mutate, OfficeProtectionCoverageState remove,
        string api, string limitation) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Capability id cannot be empty.", nameof(id));
        if (string.IsNullOrWhiteSpace(formatId)) throw new ArgumentException("Format id cannot be empty.", nameof(formatId));
        if (string.IsNullOrWhiteSpace(packageId)) throw new ArgumentException("Package id cannot be empty.", nameof(packageId));
        if (string.IsNullOrWhiteSpace(api)) throw new ArgumentException("Capability API cannot be empty.", nameof(api));
        Id = id.Trim();
        FormatId = formatId.Trim();
        PackageId = packageId.Trim();
        Kind = kind;
        Inspect = inspect;
        Open = open;
        Create = create;
        Validate = validate;
        Mutate = mutate;
        Remove = remove;
        Api = api.Trim();
        Limitation = limitation?.Trim() ?? string.Empty;
    }

    /// <summary>Stable capability identifier.</summary>
    public string Id { get; }
    /// <summary>Format or format family.</summary>
    public string FormatId { get; }
    /// <summary>Package that owns the format behavior.</summary>
    public string PackageId { get; }
    /// <summary>Protection mechanism classification.</summary>
    public OfficeProtectionKind Kind { get; }
    /// <summary>Inspection coverage.</summary>
    public OfficeProtectionCoverageState Inspect { get; }
    /// <summary>Decryption, deobfuscation, or protected-open coverage.</summary>
    public OfficeProtectionCoverageState Open { get; }
    /// <summary>Protection or signature creation coverage.</summary>
    public OfficeProtectionCoverageState Create { get; }
    /// <summary>Cryptographic validation coverage.</summary>
    public OfficeProtectionCoverageState Validate { get; }
    /// <summary>Safe mutation or round-trip coverage.</summary>
    public OfficeProtectionCoverageState Mutate { get; }
    /// <summary>Explicit protection-removal coverage.</summary>
    public OfficeProtectionCoverageState Remove { get; }
    /// <summary>Primary public API entry point.</summary>
    public string Api { get; }
    /// <summary>Important boundary or limitation.</summary>
    public string Limitation { get; }
}
