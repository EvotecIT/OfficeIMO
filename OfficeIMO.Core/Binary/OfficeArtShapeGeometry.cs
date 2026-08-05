using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

/// <summary>
/// Decodes shape-specific OfficeArt geometry values without assuming that adjustment units are
/// interchangeable between preset shape families.
/// </summary>
public sealed class OfficeArtShapeGeometry {
    private const ushort FirstAdjustmentPropertyId = 0x0147;
    private const int AdjustmentCount = 8;

    private OfficeArtShapeGeometry(IReadOnlyList<int?> adjustmentValues) {
        AdjustmentValues = new ReadOnlyCollection<int?>(adjustmentValues.ToArray());
    }

    /// <summary>Decodes geometry adjustment values from an OfficeArt property table.</summary>
    public static OfficeArtShapeGeometry Decode(IReadOnlyList<OfficeArtProperty>? properties) {
        IReadOnlyList<OfficeArtProperty> source = properties ?? Array.Empty<OfficeArtProperty>();
        var adjustments = new int?[AdjustmentCount];
        for (int index = 0; index < adjustments.Length; index++) {
            ushort propertyId = checked((ushort)(FirstAdjustmentPropertyId + index));
            OfficeArtProperty? property = source.LastOrDefault(item =>
                item.PropertyId == propertyId && !item.IsComplex);
            if (property != null) adjustments[index] = unchecked((int)property.Value);
        }
        return new OfficeArtShapeGeometry(adjustments);
    }

    /// <summary>
    /// Gets the eight optional signed adjustment slots corresponding to adjustValue through
    /// adjust8Value. Their interpretation is shape-specific.
    /// </summary>
    public IReadOnlyList<int?> AdjustmentValues { get; }

    /// <summary>Gets whether at least one explicit adjustment value is present.</summary>
    public bool HasAdjustments => AdjustmentValues.Any(value => value.HasValue);
}
