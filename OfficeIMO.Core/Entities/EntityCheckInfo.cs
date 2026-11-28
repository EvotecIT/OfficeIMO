using System;

namespace OfficeIMO.Entities;

/// <summary>Identifies an available checksum and the documentation used to implement it.</summary>
public sealed class EntityCheckInfo {
    /// <summary>Creates immutable checksum metadata for a scheme or one of its variants.</summary>
    public EntityCheckInfo(string name, EntityCheckKind kind, EntityCheckProvenance provenance, Uri? source = null) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A checksum name is required.", nameof(name));
        Name = name;
        Kind = kind;
        Provenance = provenance;
        Source = source;
    }

    /// <summary>Scheme-defined name of the check.</summary>
    public string Name { get; }
    /// <summary>Algorithm family.</summary>
    public EntityCheckKind Kind { get; }
    /// <summary>Publication status of the algorithm.</summary>
    public EntityCheckProvenance Provenance { get; }
    /// <summary>Publication that describes the algorithm.</summary>
    public Uri? Source { get; }
}
