using System;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Word line spacing expressed as a line multiple or a physical line height.</summary>
/// <remarks>Default is an explicit zero multiple. Native quantization occurs when assigned to a paragraph.</remarks>
public readonly struct WordLineSpacing : IEquatable<WordLineSpacing> {
    private readonly double _multiple;
    private readonly global::OfficeIMO.OfficeLength _length;
    private WordLineSpacing(WordLineSpacingRule rule, double multiple, global::OfficeIMO.OfficeLength length) {
        Rule = rule; _multiple = multiple; _length = length;
    }
    /// <summary>Gets the spacing interpretation; it cannot be changed independently of the value.</summary>
    public WordLineSpacingRule Rule { get; }
    /// <summary>Gets a line multiple for automatic spacing, otherwise null.</summary>
    public double? MultipleValue => Rule == WordLineSpacingRule.Auto ? _multiple : (double?)null;
    /// <summary>Gets a physical height for exact or minimum spacing, otherwise null.</summary>
    public global::OfficeIMO.OfficeLength? Height => Rule == WordLineSpacingRule.Auto ? null : _length;
    /// <summary>Creates automatic spacing. One means single spacing; zero remains an explicit authored zero.</summary>
    public static WordLineSpacing Multiple(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0 || value > int.MaxValue / 240d)
            throw new ArgumentOutOfRangeException(nameof(value), "The multiple must fit Word's nonnegative line-spacing range.");
        return new WordLineSpacing(WordLineSpacingRule.Auto, value, default);
    }
    /// <summary>Creates a fixed physical line height; zero remains explicit.</summary>
    public static WordLineSpacing Exact(global::OfficeIMO.OfficeLength height) => Physical(WordLineSpacingRule.Exact, height);
    /// <summary>Creates a minimum physical line height; zero remains explicit.</summary>
    public static WordLineSpacing AtLeast(global::OfficeIMO.OfficeLength height) => Physical(WordLineSpacingRule.AtLeast, height);
    private static WordLineSpacing Physical(WordLineSpacingRule rule, global::OfficeIMO.OfficeLength height) {
        if (height < global::OfficeIMO.OfficeLength.Zero || height.ToTwips() > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(height), "The height must fit Word's nonnegative line-spacing range.");
        return new WordLineSpacing(rule, 0, height);
    }
    internal int ToNativeValue() => Rule == WordLineSpacingRule.Auto
        ? checked((int)Math.Round(_multiple * 240d, MidpointRounding.AwayFromZero))
        : checked((int)_length.ToTwips());
    internal static WordLineSpacing? Read(SpacingBetweenLines? spacing) {
        if (!int.TryParse(spacing?.Line?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int line) || line < 0)
            return null;
        // Read the lexical token: accessing EnumValue.Value can throw on a future/unknown token.
        string? rule = spacing?.LineRule?.InnerText;
        if (rule == null || rule == "auto") return Multiple(line / 240d);
        var height = global::OfficeIMO.OfficeLength.Emus(checked((long)line * 635));
        if (rule == "exact") return Exact(height);
        if (rule == "atLeast") return AtLeast(height);
        return null;
    }
    /// <inheritdoc />
    public bool Equals(WordLineSpacing other) => Rule == other.Rule && _multiple.Equals(other._multiple) && _length == other._length;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is WordLineSpacing other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked(((int)Rule * 397 ^ _multiple.GetHashCode()) * 397 ^ _length.GetHashCode());
    /// <summary>Compares spacing values, including their interpretation.</summary>
    public static bool operator ==(WordLineSpacing left, WordLineSpacing right) => left.Equals(right);
    /// <summary>Compares spacing values, including their interpretation.</summary>
    public static bool operator !=(WordLineSpacing left, WordLineSpacing right) => !left.Equals(right);
}
