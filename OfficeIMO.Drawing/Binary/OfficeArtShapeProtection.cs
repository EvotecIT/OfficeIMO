using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

/// <summary>
/// Decodes the shared OfficeArt protection Boolean property for a shape.
/// </summary>
public sealed class OfficeArtShapeProtection {
    private OfficeArtShapeProtection(
        IReadOnlyList<OfficeArtProperty> properties) {
        Properties = properties?.ToArray() ?? Array.Empty<OfficeArtProperty>();
        LockAgainstUngrouping = GetBoolean(6);
        LockRotation = GetBoolean(7);
        LockAspectRatio = GetBoolean(8);
        LockPosition = GetBoolean(9);
        LockAgainstSelect = GetBoolean(10);
        LockCropping = GetBoolean(11);
        LockVertices = GetBoolean(12);
        LockText = GetBoolean(13);
        LockAdjustHandles = GetBoolean(14);
        LockAgainstGrouping = GetBoolean(15);
    }

    /// <summary>Decodes shape-protection state from an OfficeArt property table.</summary>
    public static OfficeArtShapeProtection Decode(
        IReadOnlyList<OfficeArtProperty>? properties) =>
        new(properties ?? Array.Empty<OfficeArtProperty>());

    /// <summary>Gets the source property entries.</summary>
    public IReadOnlyList<OfficeArtProperty> Properties { get; }

    /// <summary>Gets whether a grouped shape is explicitly locked against ungrouping.</summary>
    public bool? LockAgainstUngrouping { get; }

    /// <summary>Gets whether rotation is explicitly locked.</summary>
    public bool? LockRotation { get; }

    /// <summary>Gets whether aspect-ratio changes are explicitly locked.</summary>
    public bool? LockAspectRatio { get; }

    /// <summary>Gets whether position changes are explicitly locked.</summary>
    public bool? LockPosition { get; }

    /// <summary>Gets whether selection is explicitly locked.</summary>
    public bool? LockAgainstSelect { get; }

    /// <summary>Gets whether picture cropping is explicitly locked.</summary>
    public bool? LockCropping { get; }

    /// <summary>Gets whether path-vertex editing is explicitly locked.</summary>
    public bool? LockVertices { get; }

    /// <summary>Gets whether attached text editing is explicitly locked.</summary>
    public bool? LockText { get; }

    /// <summary>Gets whether geometry-adjustment handles are explicitly locked.</summary>
    public bool? LockAdjustHandles { get; }

    /// <summary>Gets whether grouping with other shapes is explicitly locked.</summary>
    public bool? LockAgainstGrouping { get; }

    private bool? GetBoolean(int useBit) {
        OfficeArtProperty? property = Properties.LastOrDefault(item =>
            item.PropertyId == 0x007F && !item.IsComplex);
        if (property == null || (property.Value & (1U << useBit)) == 0) {
            return null;
        }
        return (property.Value & (1U << checked(useBit + 16))) != 0;
    }
}
