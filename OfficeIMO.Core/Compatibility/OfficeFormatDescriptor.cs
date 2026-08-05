using System;

namespace OfficeIMO.Drawing;

/// <summary>Describes one concrete Office file format and document kind.</summary>
public sealed class OfficeFormatDescriptor : IEquatable<OfficeFormatDescriptor> {
    /// <summary>Creates a format descriptor.</summary>
    public OfficeFormatDescriptor(
        string id,
        string extension,
        OfficeDocumentFamily family,
        OfficeDocumentKind documentKind,
        OfficeFormatGeneration generation,
        OfficeFormatEncoding encoding,
        bool macroEnabled) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Format id cannot be empty.", nameof(id));
        if (string.IsNullOrWhiteSpace(extension)) throw new ArgumentException("Format extension cannot be empty.", nameof(extension));

        Id = id.Trim();
        Extension = NormalizeExtension(extension);
        Family = family;
        DocumentKind = documentKind;
        Generation = generation;
        Encoding = encoding;
        IsMacroEnabled = macroEnabled;
    }

    /// <summary>Gets the stable format identifier, such as <c>Word.Docx</c>.</summary>
    public string Id { get; }

    /// <summary>Gets the normalized lower-case extension including its leading period.</summary>
    public string Extension { get; }

    /// <summary>Gets the owning Office document family.</summary>
    public OfficeDocumentFamily Family { get; }

    /// <summary>Gets the logical document kind.</summary>
    public OfficeDocumentKind DocumentKind { get; }

    /// <summary>Gets the format generation.</summary>
    public OfficeFormatGeneration Generation { get; }

    /// <summary>Gets the physical encoding.</summary>
    public OfficeFormatEncoding Encoding { get; }

    /// <summary>Gets whether the format can carry VBA projects.</summary>
    public bool IsMacroEnabled { get; }

    /// <inheritdoc />
    public bool Equals(OfficeFormatDescriptor? other) => other != null
        && string.Equals(Id, other.Id, StringComparison.Ordinal)
        && string.Equals(Extension, other.Extension, StringComparison.Ordinal);

    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as OfficeFormatDescriptor);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            return (StringComparer.Ordinal.GetHashCode(Id) * 397)
                ^ StringComparer.Ordinal.GetHashCode(Extension);
        }
    }

    /// <inheritdoc />
    public override string ToString() => $"{Id} ({Extension})";

    private static string NormalizeExtension(string extension) {
        string value = extension.Trim();
        if (!value.StartsWith(".", StringComparison.Ordinal)) value = "." + value;
        return value.ToLowerInvariant();
    }
}
