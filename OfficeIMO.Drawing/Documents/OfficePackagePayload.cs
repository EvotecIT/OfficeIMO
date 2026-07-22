using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Classifies an executable or embedded payload stored in an Office Open XML package.</summary>
public enum OfficeEmbeddedPayloadKind {
    /// <summary>An embedded document or arbitrary package payload.</summary>
    EmbeddedPackage,
    /// <summary>An OLE object payload.</summary>
    OleObject,
    /// <summary>An ActiveX control definition or binary persistence payload.</summary>
    ActiveX,
    /// <summary>Another payload stored below an Office embeddings package path.</summary>
    Other
}

/// <summary>Describes an embedded package payload and the relationship that owns it.</summary>
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
        string? sha256) {
        Id = id;
        Kind = kind;
        OwnerPartUri = ownerPartUri;
        RelationshipId = relationshipId;
        PartUri = partUri;
        ContentType = contentType;
        SuggestedFileName = suggestedFileName;
        Length = length;
        Sha256 = sha256;
    }

    /// <summary>Stable package-local id composed from the owner part URI and relationship id.</summary>
    public string Id { get; }

    /// <summary>Payload classification.</summary>
    public OfficeEmbeddedPayloadKind Kind { get; }

    /// <summary>URI of the package part that owns the relationship.</summary>
    public string OwnerPartUri { get; }

    /// <summary>Relationship id used by the owner part.</summary>
    public string RelationshipId { get; }

    /// <summary>URI of the binary or embedded package part.</summary>
    public string PartUri { get; }

    /// <summary>Declared MIME content type.</summary>
    public string ContentType { get; }

    /// <summary>Best-effort file name inferred from the package part URI.</summary>
    public string SuggestedFileName { get; }

    /// <summary>Uncompressed payload length in bytes.</summary>
    public long Length { get; }

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
        bool hasSignature) {
        PartUri = partUri;
        ContentType = contentType;
        Length = length;
        Sha256 = sha256;
        ModuleNames = moduleNames;
        HasSignature = hasSignature;
    }

    /// <summary>URI of the <c>vbaProject.bin</c> package part.</summary>
    public string PartUri { get; }

    /// <summary>Declared MIME content type.</summary>
    public string ContentType { get; }

    /// <summary>VBA project length in bytes.</summary>
    public long Length { get; }

    /// <summary>Lowercase SHA-256 digest when hash calculation was requested; otherwise null.</summary>
    public string? Sha256 { get; }

    /// <summary>Best-effort VBA module names discovered in the compound project.</summary>
    public IReadOnlyList<string> ModuleNames { get; }

    /// <summary>True when the VBA project has a signature relationship.</summary>
    public bool HasSignature { get; }
}
