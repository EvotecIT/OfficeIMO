using System;
using System.Collections.Generic;

namespace OfficeIMO;

/// <summary>Classifies an inert embedded payload identified by a native OfficeIMO owner.</summary>
public enum OfficeEmbeddedPayloadKind {
    /// <summary>An embedded document or arbitrary package payload.</summary>
    EmbeddedPackage,
    /// <summary>An OLE object payload.</summary>
    OleObject,
    /// <summary>An ActiveX control definition or binary persistence payload.</summary>
    ActiveX,
    /// <summary>Another native embedded or binary payload.</summary>
    Other
}

/// <summary>Describes a native embedded payload and its part, relationship, or source-profile identity.</summary>
public sealed class OfficeEmbeddedPayloadInfo {
    internal OfficeEmbeddedPayloadInfo(
        string id,
        OfficeEmbeddedPayloadKind kind,
        string ownerPartUri,
        string relationshipId,
        string partUri,
        string contentType,
        string suggestedFileName,
        long length,
        string? sha256,
        int ordinal = -1) {
        Id = id;
        Kind = kind;
        OwnerPartUri = ownerPartUri;
        RelationshipId = relationshipId;
        PartUri = partUri;
        ContentType = contentType;
        SuggestedFileName = suggestedFileName;
        Length = length;
        Sha256 = sha256;
        Ordinal = ordinal;
    }

    /// <summary>Stable native payload identity. Open XML uses owner part URI and relationship id; other owners define their source-profile identity.</summary>
    public string Id { get; }

    /// <summary>Payload classification.</summary>
    public OfficeEmbeddedPayloadKind Kind { get; }

    /// <summary>Native owner part or source-profile locator. It is not a filesystem path or physical byte offset.</summary>
    public string OwnerPartUri { get; }

    /// <summary>Relationship id used by the owner part, or empty when the source profile has no relationship model.</summary>
    public string RelationshipId { get; }

    /// <summary>Native payload part or source-profile locator. It is not a physical file offset.</summary>
    public string PartUri { get; }

    /// <summary>Declared MIME content type.</summary>
    public string ContentType { get; }

    /// <summary>Best-effort file name inferred from the package part URI.</summary>
    public string SuggestedFileName { get; }

    /// <summary>Decoded payload length in bytes, or -1 when the native owner has no size metadata without reading the payload.</summary>
    public long Length { get; }

    /// <summary>Whether <see cref="Length"/> is known from metadata or a requested bounded read.</summary>
    public bool HasKnownLength => Length >= 0;

    /// <summary>Zero-based inventory ordinal in an inspection session, or -1 for a standalone compatibility descriptor.</summary>
    public int Ordinal { get; }

    /// <summary>Lowercase SHA-256 digest when hash calculation was requested; otherwise null.</summary>
    public string? Sha256 { get; }
}

/// <summary>Describes a VBA project embedded in an Office document.</summary>
public sealed class OfficeVbaProjectInfo {
    /// <summary>Default maximum VBA project bytes materialized by metadata inspection.</summary>
    public const long DefaultMaximumProjectBytes = 64L * 1024L * 1024L;

    internal OfficeVbaProjectInfo(
        string partUri,
        string contentType,
        long length,
        string? sha256,
        IReadOnlyList<string> moduleNames,
        bool hasSignature,
        bool moduleInspectionRequested = true) {
        PartUri = partUri;
        ContentType = contentType;
        Length = length;
        Sha256 = sha256;
        ModuleNames = moduleNames;
        HasSignature = hasSignature;
        ModuleInspectionRequested = moduleInspectionRequested;
    }

    /// <summary>URI of the <c>vbaProject.bin</c> package part.</summary>
    public string PartUri { get; }

    /// <summary>Declared MIME content type.</summary>
    public string ContentType { get; }

    /// <summary>VBA project length in bytes, or -1 for a descriptor that did not request a payload read.</summary>
    public long Length { get; }

    /// <summary>Whether the decoded project length is known.</summary>
    public bool HasKnownLength => Length >= 0;

    /// <summary>Lowercase SHA-256 digest when hash calculation was requested; otherwise null.</summary>
    public string? Sha256 { get; }

    /// <summary>Best-effort VBA module names discovered in the compound project.</summary>
    public IReadOnlyList<string> ModuleNames { get; }

    /// <summary>Whether binary VBA module inspection was requested. False distinguishes an unrequested payload from an inspected empty module list.</summary>
    public bool ModuleInspectionRequested { get; }

    /// <summary>True when the VBA project has a signature relationship.</summary>
    public bool HasSignature { get; }
}
