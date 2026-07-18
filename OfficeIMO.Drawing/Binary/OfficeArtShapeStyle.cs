using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

/// <summary>
/// Decodes common fill, line, and shadow state from an OfficeArt property table.
/// Numeric enum values remain available so application-specific projectors can map them without loss.
/// </summary>
public sealed class OfficeArtShapeStyle {
    private OfficeArtShapeStyle(IReadOnlyList<OfficeArtProperty> properties) {
        Properties = properties?.ToArray() ?? Array.Empty<OfficeArtProperty>();
        FillType = GetUInt32(0x0180);
        FillColor = GetColor(0x0181);
        FillOpacity = GetFixedPoint(0x0182);
        FillBackColor = GetColor(0x0183);
        FillBackOpacity = GetFixedPoint(0x0184);
        FillBlipStoreIndex = GetBlipStoreIndex(0x0186);
        FillAngleDegrees = GetSignedFixedPoint(0x018B);
        FillFocusPercent = GetInt32(0x018C);
        FillGradientStops = ReadGradientStops(out bool gradientStopsTruncated);
        IsFillGradientStopTableTruncated = gradientStopsTruncated;
        FillEnabled = GetBoolean(0x01BF, 0x00100000U, 0x00000010U);
        LineColor = GetColor(0x01C0);
        LineOpacity = GetFixedPoint(0x01C1);
        LineType = GetUInt32(0x01C4);
        LineWidthEmus = GetInt32(0x01CB);
        LineStyle = GetUInt32(0x01CD);
        LineDashing = GetUInt32(0x01CE);
        LineStartArrowhead = GetUInt32(0x01D0);
        LineEndArrowhead = GetUInt32(0x01D1);
        LineStartArrowWidth = GetUInt32(0x01D2);
        LineStartArrowLength = GetUInt32(0x01D3);
        LineEndArrowWidth = GetUInt32(0x01D4);
        LineEndArrowLength = GetUInt32(0x01D5);
        LineJoinStyle = GetUInt32(0x01D6);
        LineEndCapStyle = GetUInt32(0x01D7);
        LineEnabled = GetBoolean(0x01FF, 0x00080000U, 0x00000008U);
        ShadowType = GetUInt32(0x0200);
        ShadowColor = GetColor(0x0201);
        ShadowOpacity = GetFixedPoint(0x0204);
        ShadowOffsetXEmus = GetInt32(0x0205);
        ShadowOffsetYEmus = GetInt32(0x0206);
        ShadowSoftnessEmus = GetInt32(0x021C);
        ShadowEnabled = GetBoolean(0x023F, 0x00020000U, 0x00000002U);
        PreferRelativeResize = GetBoolean(0x033F,
            1U << 11, 1U << 27);
        LockShapeType = GetBoolean(0x033F,
            1U << 12, 1U << 28);
        Hidden = GetBoolean(0x03BF,
            1U << 14, 1U << 30);
    }

    /// <summary>Decodes a shape style from OfficeArt properties.</summary>
    public static OfficeArtShapeStyle Decode(IReadOnlyList<OfficeArtProperty>? properties) =>
        new OfficeArtShapeStyle(properties ?? Array.Empty<OfficeArtProperty>());

    /// <summary>Gets the source property entries.</summary>
    public IReadOnlyList<OfficeArtProperty> Properties { get; }

    /// <summary>Gets the MSOFILLTYPE value.</summary>
    public uint? FillType { get; }

    /// <summary>Gets the fill color reference.</summary>
    public OfficeArtColorReference? FillColor { get; }

    /// <summary>Gets fill opacity from 0 through 1.</summary>
    public double? FillOpacity { get; }

    /// <summary>Gets the fill background color reference used by patterns and gradients.</summary>
    public OfficeArtColorReference? FillBackColor { get; }

    /// <summary>Gets fill-background opacity from 0 through 1.</summary>
    public double? FillBackOpacity { get; }

    /// <summary>Gets the one-based BLIP store index used by picture, texture, or pattern fills.</summary>
    public int? FillBlipStoreIndex { get; }

    /// <summary>Gets the signed 16.16 gradient angle in counterclockwise degrees.</summary>
    public double? FillAngleDegrees { get; }

    /// <summary>Gets the gradient focus position from -100 through 100.</summary>
    public int? FillFocusPercent { get; }

    /// <summary>Gets decoded MSOSHADECOLOR gradient stops in source order.</summary>
    public IReadOnlyList<OfficeArtGradientStop> FillGradientStops { get; }

    /// <summary>Gets whether a declared gradient-stop array is malformed or truncated.</summary>
    public bool IsFillGradientStopTableTruncated { get; }

    /// <summary>Gets explicit fill visibility, or null when the property inherits its default.</summary>
    public bool? FillEnabled { get; }

    /// <summary>Gets the line color reference.</summary>
    public OfficeArtColorReference? LineColor { get; }

    /// <summary>Gets line opacity from 0 through 1.</summary>
    public double? LineOpacity { get; }

    /// <summary>Gets the MSOLINETYPE value.</summary>
    public uint? LineType { get; }

    /// <summary>Gets the line width in English Metric Units.</summary>
    public int? LineWidthEmus { get; }

    /// <summary>Gets the MSOLINESTYLE value.</summary>
    public uint? LineStyle { get; }

    /// <summary>Gets the MSOLINEDASHING value.</summary>
    public uint? LineDashing { get; }

    /// <summary>Gets the start MSOLINEEND value.</summary>
    public uint? LineStartArrowhead { get; }

    /// <summary>Gets the end MSOLINEEND value.</summary>
    public uint? LineEndArrowhead { get; }

    /// <summary>Gets the start MSOLINEENDWIDTH value.</summary>
    public uint? LineStartArrowWidth { get; }

    /// <summary>Gets the start MSOLINEENDLENGTH value.</summary>
    public uint? LineStartArrowLength { get; }

    /// <summary>Gets the end MSOLINEENDWIDTH value.</summary>
    public uint? LineEndArrowWidth { get; }

    /// <summary>Gets the end MSOLINEENDLENGTH value.</summary>
    public uint? LineEndArrowLength { get; }

    /// <summary>Gets the MSOLINEJOIN value.</summary>
    public uint? LineJoinStyle { get; }

    /// <summary>Gets the MSOLINECAP value.</summary>
    public uint? LineEndCapStyle { get; }

    /// <summary>Gets explicit line visibility, or null when the property inherits its default.</summary>
    public bool? LineEnabled { get; }

    /// <summary>Gets the MSOSHADOWTYPE value.</summary>
    public uint? ShadowType { get; }

    /// <summary>Gets the primary shadow color reference.</summary>
    public OfficeArtColorReference? ShadowColor { get; }

    /// <summary>Gets shadow opacity from 0 through 1.</summary>
    public double? ShadowOpacity { get; }

    /// <summary>Gets the signed horizontal shadow offset in English Metric Units.</summary>
    public int? ShadowOffsetXEmus { get; }

    /// <summary>Gets the signed vertical shadow offset in English Metric Units.</summary>
    public int? ShadowOffsetYEmus { get; }

    /// <summary>Gets the shadow blur radius in English Metric Units.</summary>
    public int? ShadowSoftnessEmus { get; }

    /// <summary>Gets explicit shadow visibility, or null when the property inherits its default.</summary>
    public bool? ShadowEnabled { get; }

    /// <summary>
    /// Gets whether the resizing user interface explicitly prefers values
    /// relative to the original shape size.
    /// </summary>
    public bool? PreferRelativeResize { get; }

    /// <summary>Gets whether changing the shape type is explicitly locked.</summary>
    public bool? LockShapeType { get; }

    /// <summary>Gets whether the shape is explicitly hidden from display.</summary>
    public bool? Hidden { get; }

    /// <summary>
    /// Gets whether the hidden-state bits can be rewritten while preserving
    /// every unrelated Group Shape Boolean property.
    /// </summary>
    public bool CanRewriteHiddenState => Properties
        .Where(property => property.PropertyId == 0x03BF)
        .All(property => !property.IsComplex && !property.IsBlipId);

    /// <summary>Gets whether this style includes fill or line values that can be projected directly.</summary>
    public bool HasProjectableStyle => FillEnabled.HasValue || FillColor.HasValue || FillOpacity.HasValue
        || FillBackColor.HasValue || FillBackOpacity.HasValue || FillBlipStoreIndex.HasValue
        || FillAngleDegrees.HasValue || FillFocusPercent.HasValue
        || LineEnabled.HasValue || LineColor.HasValue || LineOpacity.HasValue || LineWidthEmus.HasValue
        || LineDashing.HasValue || LineStartArrowhead.HasValue || LineEndArrowhead.HasValue
        || LineJoinStyle.HasValue || LineEndCapStyle.HasValue || HasProjectableShadow;

    /// <summary>Gets whether an enabled offset shadow can be projected directly.</summary>
    public bool HasProjectableShadow => ShadowEnabled == true && ShadowType.GetValueOrDefault() == 0;

    /// <summary>Gets whether enabled visual state requires a richer projector than solid fills and lines.</summary>
    public bool HasUnprojectedVisualStyle =>
        FillEnabled != false && FillType.GetValueOrDefault() > 0
        || LineEnabled != false && (LineType.GetValueOrDefault() > 0 || LineStyle.GetValueOrDefault() > 0)
        || LineEnabled != false && Properties.Any(property => property.PropertyId == 0x01CF)
        || ShadowEnabled == true && ShadowType.GetValueOrDefault() != 0;

    /// <summary>
    /// Gets whether every source fill, line, and shadow property is owned by
    /// the editable solid-style projection and can therefore be rewritten
    /// without discarding an opaque visual property.
    /// </summary>
    public bool CanRewriteProjectedVisualStyle =>
        !HasUnprojectedVisualStyle
        && Properties.All(IsRewritableVisualProperty);

    private static bool IsRewritableVisualProperty(
        OfficeArtProperty property) {
        if (property.PropertyId is not (>= 0x0180 and <= 0x023F)) {
            return true;
        }
        if (property.IsComplex || property.IsBlipId) return false;
        return property.PropertyId is 0x0180 or 0x0181 or 0x0182
            or 0x01BF
            or 0x01C0 or 0x01C1 or 0x01CB or 0x01CE
            or >= 0x01D0 and <= 0x01D7
            or 0x01FF
            or 0x0200 or 0x0201 or 0x0204 or 0x0205 or 0x0206
            or 0x021C or 0x023F;
    }

    private OfficeArtProperty? GetProperty(ushort propertyId) =>
        Properties.LastOrDefault(property => property.PropertyId == propertyId && !property.IsComplex);

    private uint? GetUInt32(ushort propertyId) => GetProperty(propertyId)?.Value;

    private int? GetInt32(ushort propertyId) {
        OfficeArtProperty? property = GetProperty(propertyId);
        return property == null ? null : unchecked((int)property.Value);
    }

    private OfficeArtColorReference? GetColor(ushort propertyId) {
        uint? value = GetUInt32(propertyId);
        return value.HasValue ? new OfficeArtColorReference(value.Value) : null;
    }

    private double? GetFixedPoint(ushort propertyId) {
        uint? value = GetUInt32(propertyId);
        if (!value.HasValue || value.Value > 0x00010000U) return null;
        return value.Value / 65536D;
    }

    private double? GetSignedFixedPoint(ushort propertyId) {
        uint? value = GetUInt32(propertyId);
        return value.HasValue ? unchecked((int)value.Value) / 65536D : null;
    }

    private int? GetBlipStoreIndex(ushort propertyId) {
        OfficeArtProperty? property = Properties.LastOrDefault(candidate =>
            candidate.PropertyId == propertyId && candidate.IsBlipId && !candidate.IsComplex);
        return property == null || property.Value == 0 || property.Value > int.MaxValue
            ? null
            : unchecked((int)property.Value);
    }

    private IReadOnlyList<OfficeArtGradientStop> ReadGradientStops(out bool truncated) {
        truncated = false;
        OfficeArtProperty? property = Properties.LastOrDefault(candidate =>
            candidate.PropertyId == 0x0197 && candidate.IsComplex);
        if (property == null) return Array.Empty<OfficeArtGradientStop>();
        byte[]? data = property.CopyComplexData();
        if (data == null || data.Length < 6) {
            truncated = true;
            return Array.Empty<OfficeArtGradientStop>();
        }
        int count = ReadUInt16(data, 0);
        int allocated = ReadUInt16(data, 2);
        int elementSize = ReadUInt16(data, 4);
        if (count > allocated || elementSize != 8 || count > (data.Length - 6) / 8) {
            truncated = true;
            return Array.Empty<OfficeArtGradientStop>();
        }

        var result = new List<OfficeArtGradientStop>(count);
        double previous = -1D;
        for (int index = 0; index < count; index++) {
            int offset = 6 + index * 8;
            uint rawPosition = ReadUInt32(data, offset + 4);
            if (rawPosition > 0x00010000U) {
                truncated = true;
                return Array.Empty<OfficeArtGradientStop>();
            }
            double position = rawPosition / 65536D;
            if (position < previous) {
                truncated = true;
                return Array.Empty<OfficeArtGradientStop>();
            }
            result.Add(new OfficeArtGradientStop(
                new OfficeArtColorReference(ReadUInt32(data, offset)), position));
            previous = position;
        }
        return result;
    }

    private static ushort ReadUInt16(byte[] data, int offset) => unchecked((ushort)(
        data[offset] | data[offset + 1] << 8));

    private static uint ReadUInt32(byte[] data, int offset) => unchecked((uint)(
        data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16
        | data[offset + 3] << 24));

    private bool? GetBoolean(ushort propertyId, uint useMask, uint valueMask) {
        uint? value = GetUInt32(propertyId);
        if (!value.HasValue || (value.Value & useMask) == 0) return null;
        return (value.Value & valueMask) != 0;
    }
}
