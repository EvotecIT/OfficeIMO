using System;

namespace OfficeIMO.Drawing;

/// <summary>CSS-compatible face attributes used when selecting among registered font faces.</summary>
public readonly struct OfficeFontFaceDescriptor : IEquatable<OfficeFontFaceDescriptor> {
    private readonly int _weight;
    private readonly double _stretchPercent;
    private readonly double _obliqueAngleDegrees;

    /// <summary>Creates a face descriptor.</summary>
    /// <param name="weight">Numeric font weight from 1 through 1000.</param>
    /// <param name="stretchPercent">Face width as a percentage of normal width.</param>
    /// <param name="slant">Normal, italic, or oblique face classification.</param>
    /// <param name="obliqueAngleDegrees">Oblique angle in degrees when <paramref name="slant"/> is oblique.</param>
    public OfficeFontFaceDescriptor(
        int weight = 400,
        double stretchPercent = 100D,
        OfficeFontSlant slant = OfficeFontSlant.Normal,
        double obliqueAngleDegrees = 14D) {
        if (weight < 1 || weight > 1000) throw new ArgumentOutOfRangeException(nameof(weight));
        if (double.IsNaN(stretchPercent) || double.IsInfinity(stretchPercent) || stretchPercent < 50D || stretchPercent > 200D) {
            throw new ArgumentOutOfRangeException(nameof(stretchPercent), "Font stretch must be between 50% and 200%.");
        }
        if (slant < OfficeFontSlant.Normal || slant > OfficeFontSlant.Oblique) throw new ArgumentOutOfRangeException(nameof(slant));
        if (double.IsNaN(obliqueAngleDegrees) || double.IsInfinity(obliqueAngleDegrees) || obliqueAngleDegrees <= -90D || obliqueAngleDegrees >= 90D) {
            throw new ArgumentOutOfRangeException(nameof(obliqueAngleDegrees), "An oblique angle must be greater than -90 and less than 90 degrees.");
        }

        _weight = weight;
        _stretchPercent = stretchPercent;
        Slant = slant;
        _obliqueAngleDegrees = slant == OfficeFontSlant.Oblique ? obliqueAngleDegrees : 0D;
    }

    /// <summary>Numeric face weight from 1 through 1000.</summary>
    public int Weight => _weight == 0 ? 400 : _weight;

    /// <summary>Face width as a percentage of normal width.</summary>
    public double StretchPercent => _stretchPercent == 0D ? 100D : _stretchPercent;

    /// <summary>Normal, italic, or oblique face classification.</summary>
    public OfficeFontSlant Slant { get; }

    /// <summary>Oblique face angle in degrees, or zero for normal and italic faces.</summary>
    public double ObliqueAngleDegrees => Slant == OfficeFontSlant.Oblique ? _obliqueAngleDegrees : 0D;

    /// <summary>Regular 400-weight, normal-width, upright face.</summary>
    public static OfficeFontFaceDescriptor Regular => new OfficeFontFaceDescriptor();

    /// <summary>Creates a descriptor matching the legacy bold/italic flags.</summary>
    public static OfficeFontFaceDescriptor FromStyle(OfficeFontStyle style) => new OfficeFontFaceDescriptor(
        (style & OfficeFontStyle.Bold) == OfficeFontStyle.Bold ? 700 : 400,
        100D,
        (style & OfficeFontStyle.Italic) == OfficeFontStyle.Italic ? OfficeFontSlant.Italic : OfficeFontSlant.Normal);

    /// <summary>Compatibility style flags inferred from the numeric face attributes.</summary>
    public OfficeFontStyle ToStyle() {
        OfficeFontStyle style = Weight >= 600 ? OfficeFontStyle.Bold : OfficeFontStyle.Regular;
        if (Slant != OfficeFontSlant.Normal) style |= OfficeFontStyle.Italic;
        return style;
    }

    /// <inheritdoc />
    public bool Equals(OfficeFontFaceDescriptor other) =>
        Weight == other.Weight &&
        StretchPercent.Equals(other.StretchPercent) &&
        Slant == other.Slant &&
        ObliqueAngleDegrees.Equals(other.ObliqueAngleDegrees);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeFontFaceDescriptor other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = Weight;
            hash = (hash * 397) ^ StretchPercent.GetHashCode();
            hash = (hash * 397) ^ (int)Slant;
            hash = (hash * 397) ^ ObliqueAngleDegrees.GetHashCode();
            return hash;
        }
    }

    /// <summary>Equality operator.</summary>
    public static bool operator ==(OfficeFontFaceDescriptor left, OfficeFontFaceDescriptor right) => left.Equals(right);

    /// <summary>Inequality operator.</summary>
    public static bool operator !=(OfficeFontFaceDescriptor left, OfficeFontFaceDescriptor right) => !left.Equals(right);
}

/// <summary>Font face slant classification.</summary>
public enum OfficeFontSlant {
    /// <summary>Upright face.</summary>
    Normal = 0,

    /// <summary>Designed italic face.</summary>
    Italic = 1,

    /// <summary>Slanted or oblique face.</summary>
    Oblique = 2
}
