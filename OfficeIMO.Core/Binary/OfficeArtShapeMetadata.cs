using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Decodes user-facing name and description metadata from OfficeArt properties.</summary>
public sealed class OfficeArtShapeMetadata {
    private OfficeArtShapeMetadata(IReadOnlyList<OfficeArtProperty> properties) {
        Properties = properties?.ToArray() ?? Array.Empty<OfficeArtProperty>();
        Name = GetText(0x0380);
        Description = GetText(0x0381);
    }

    /// <summary>Decodes shape metadata from an OfficeArt property table.</summary>
    public static OfficeArtShapeMetadata Decode(IReadOnlyList<OfficeArtProperty>? properties) =>
        new OfficeArtShapeMetadata(properties ?? Array.Empty<OfficeArtProperty>());

    /// <summary>Gets the source property entries.</summary>
    public IReadOnlyList<OfficeArtProperty> Properties { get; }

    /// <summary>Gets the authored object name.</summary>
    public string? Name { get; }

    /// <summary>Gets the authored object description or alternative text.</summary>
    public string? Description { get; }

    /// <summary>Gets whether either user-facing metadata field is present.</summary>
    public bool HasMetadata => Name != null || Description != null;

    /// <summary>
    /// Gets whether every name or description property has complete, well-formed
    /// complex data and can therefore be replaced without masking truncation.
    /// </summary>
    public bool CanRewrite => Properties.Where(property =>
            property.PropertyId is 0x0380 or 0x0381)
        .All(property => property.IsComplex
            && property.DeclaredComplexDataLength.HasValue
            && property.AvailableComplexDataLength.HasValue
            && property.DeclaredComplexDataLength.Value
                == unchecked((uint)property.AvailableComplexDataLength.Value));

    private string? GetText(ushort propertyId) => Properties.LastOrDefault(property =>
        property.PropertyId == propertyId && property.IsComplex
        && !string.IsNullOrWhiteSpace(property.ComplexText))?.ComplexText;
}
