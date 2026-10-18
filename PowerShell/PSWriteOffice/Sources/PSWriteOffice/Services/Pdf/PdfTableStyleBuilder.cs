using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

internal sealed class PdfTableStyleOptions
{
    internal string? TableStyle { get; set; }
    internal string? HeaderFill { get; set; }
    internal string? HeaderTextColor { get; set; }
    internal string? TextColor { get; set; }
    internal string? RowStripeFill { get; set; }
    internal string? BorderColor { get; set; }
    internal double? BorderWidth { get; set; }
    internal double? FontSize { get; set; }
    internal double? HeaderFontSize { get; set; }
    internal double? LineHeight { get; set; }
    internal double? CellPaddingX { get; set; }
    internal double? CellPaddingY { get; set; }
    internal double? CellPaddingLeft { get; set; }
    internal double? CellPaddingRight { get; set; }
    internal double? CellPaddingTop { get; set; }
    internal double? CellPaddingBottom { get; set; }
    internal double? SpacingBefore { get; set; }
    internal double? SpacingAfter { get; set; }
    internal string? Caption { get; set; }
    internal PdfAlign? CaptionAlign { get; set; }
    internal string? CaptionColor { get; set; }
    internal double? CaptionFontSize { get; set; }
    internal double? CaptionSpacingAfter { get; set; }
    internal double? MaxWidth { get; set; }
    internal double? LeftIndent { get; set; }
    internal double? MinimumShrinkFontSize { get; set; }
    internal double[]? ColumnWidthPoints { get; set; }
    internal double[]? ColumnMinWidthPoints { get; set; }
    internal double[]? ColumnMaxWidthPoints { get; set; }
    internal double[]? ColumnWidthWeights { get; set; }
    internal PdfColumnAlign[]? ColumnAlign { get; set; }
    internal bool AutoFitColumns { get; set; }
    internal bool RightAlignNumeric { get; set; }
    internal bool ShrinkTextToFit { get; set; }
    internal bool KeepTogether { get; set; }
    internal bool KeepWithNext { get; set; }
    internal bool NoBorder { get; set; }
    internal bool NoHeaderFill { get; set; }
    internal bool NoRowStripeFill { get; set; }
    internal int? HeaderRowCount { get; set; }
    internal int? RepeatHeaderRowCount { get; set; }
    internal int? FooterRowCount { get; set; }

    internal bool HasAnyValue =>
        !string.IsNullOrWhiteSpace(TableStyle) ||
        !string.IsNullOrWhiteSpace(HeaderFill) ||
        !string.IsNullOrWhiteSpace(HeaderTextColor) ||
        !string.IsNullOrWhiteSpace(TextColor) ||
        !string.IsNullOrWhiteSpace(RowStripeFill) ||
        !string.IsNullOrWhiteSpace(BorderColor) ||
        BorderWidth.HasValue ||
        FontSize.HasValue ||
        HeaderFontSize.HasValue ||
        LineHeight.HasValue ||
        CellPaddingX.HasValue ||
        CellPaddingY.HasValue ||
        CellPaddingLeft.HasValue ||
        CellPaddingRight.HasValue ||
        CellPaddingTop.HasValue ||
        CellPaddingBottom.HasValue ||
        SpacingBefore.HasValue ||
        SpacingAfter.HasValue ||
        !string.IsNullOrWhiteSpace(Caption) ||
        CaptionAlign.HasValue ||
        !string.IsNullOrWhiteSpace(CaptionColor) ||
        CaptionFontSize.HasValue ||
        CaptionSpacingAfter.HasValue ||
        MaxWidth.HasValue ||
        LeftIndent.HasValue ||
        MinimumShrinkFontSize.HasValue ||
        HasValues(ColumnWidthPoints) ||
        HasValues(ColumnMinWidthPoints) ||
        HasValues(ColumnMaxWidthPoints) ||
        HasValues(ColumnWidthWeights) ||
        HasValues(ColumnAlign) ||
        AutoFitColumns ||
        RightAlignNumeric ||
        ShrinkTextToFit ||
        KeepTogether ||
        KeepWithNext ||
        NoBorder ||
        NoHeaderFill ||
        NoRowStripeFill ||
        HeaderRowCount.HasValue ||
        RepeatHeaderRowCount.HasValue ||
        FooterRowCount.HasValue;

    private static bool HasValues<T>(T[]? value) => value != null && value.Length > 0;
}

internal static class PdfTableStyleBuilder
{
    internal static PdfTableStyle? Create(PdfTableStyleOptions options)
    {
        if (!options.HasAnyValue)
        {
            return null;
        }

        var style = ResolvePreset(options.TableStyle);

        if (options.HeaderFill != null) style.HeaderFill = PdfCommandUtilities.ParseColor(options.HeaderFill);
        if (options.HeaderTextColor != null) style.HeaderTextColor = PdfCommandUtilities.ParseColor(options.HeaderTextColor);
        if (options.TextColor != null) style.TextColor = PdfCommandUtilities.ParseColor(options.TextColor);
        if (options.RowStripeFill != null) style.RowStripeFill = PdfCommandUtilities.ParseColor(options.RowStripeFill);
        if (options.BorderColor != null) style.BorderColor = PdfCommandUtilities.ParseColor(options.BorderColor);
        if (options.BorderWidth.HasValue) style.BorderWidth = options.BorderWidth.Value;
        if (options.FontSize.HasValue) style.FontSize = options.FontSize.Value;
        if (options.HeaderFontSize.HasValue) style.HeaderFontSize = options.HeaderFontSize.Value;
        if (options.LineHeight.HasValue) style.LineHeight = options.LineHeight.Value;
        if (options.CellPaddingX.HasValue) style.CellPaddingX = options.CellPaddingX.Value;
        if (options.CellPaddingY.HasValue) style.CellPaddingY = options.CellPaddingY.Value;
        if (options.CellPaddingLeft.HasValue) style.CellPaddingLeft = options.CellPaddingLeft.Value;
        if (options.CellPaddingRight.HasValue) style.CellPaddingRight = options.CellPaddingRight.Value;
        if (options.CellPaddingTop.HasValue) style.CellPaddingTop = options.CellPaddingTop.Value;
        if (options.CellPaddingBottom.HasValue) style.CellPaddingBottom = options.CellPaddingBottom.Value;
        if (options.SpacingBefore.HasValue) style.SpacingBefore = options.SpacingBefore.Value;
        if (options.SpacingAfter.HasValue) style.SpacingAfter = options.SpacingAfter.Value;
        if (options.Caption != null) style.Caption = options.Caption;
        if (options.CaptionAlign.HasValue) style.CaptionAlign = options.CaptionAlign.Value;
        if (options.CaptionColor != null) style.CaptionColor = PdfCommandUtilities.ParseColor(options.CaptionColor);
        if (options.CaptionFontSize.HasValue) style.CaptionFontSize = options.CaptionFontSize.Value;
        if (options.CaptionSpacingAfter.HasValue) style.CaptionSpacingAfter = options.CaptionSpacingAfter.Value;
        if (options.MaxWidth.HasValue) style.MaxWidth = options.MaxWidth.Value;
        if (options.LeftIndent.HasValue) style.LeftIndent = options.LeftIndent.Value;
        if (options.MinimumShrinkFontSize.HasValue) style.MinimumShrinkFontSize = options.MinimumShrinkFontSize.Value;
        if (options.ColumnWidthPoints is { Length: > 0 }) style.ColumnWidthPoints = options.ColumnWidthPoints.Select(value => (double?)value).ToList();
        if (options.ColumnMinWidthPoints is { Length: > 0 }) style.ColumnMinWidthPoints = options.ColumnMinWidthPoints.Select(value => (double?)value).ToList();
        if (options.ColumnMaxWidthPoints is { Length: > 0 }) style.ColumnMaxWidthPoints = options.ColumnMaxWidthPoints.Select(value => (double?)value).ToList();
        if (options.ColumnWidthWeights is { Length: > 0 }) style.ColumnWidthWeights = options.ColumnWidthWeights.ToList();
        if (options.ColumnAlign is { Length: > 0 }) style.Alignments = options.ColumnAlign.ToList();
        if (options.AutoFitColumns) style.AutoFitColumns = true;
        if (options.RightAlignNumeric) style.RightAlignNumeric = true;
        if (options.ShrinkTextToFit) style.ShrinkTextToFit = true;
        if (options.KeepTogether) style.KeepTogether = true;
        if (options.KeepWithNext) style.KeepWithNext = true;
        if (options.NoBorder)
        {
            style.BorderColor = null;
            style.BorderWidth = 0;
        }

        if (options.NoHeaderFill) style.HeaderFill = null;
        if (options.NoRowStripeFill) style.RowStripeFill = null;
        if (options.HeaderRowCount.HasValue) style.HeaderRowCount = options.HeaderRowCount.Value;
        if (options.RepeatHeaderRowCount.HasValue) style.RepeatHeaderRowCount = options.RepeatHeaderRowCount.Value;
        if (options.FooterRowCount.HasValue) style.FooterRowCount = options.FooterRowCount.Value;

        return style;
    }

    internal static PdfTableStyle? CreateFromSpecification(object specification)
    {
        return Create(new PdfTableStyleOptions
        {
            TableStyle = GetString(specification, "TableStyle", "Style", "WordTableStyle"),
            HeaderFill = GetString(specification, "HeaderFill"),
            HeaderTextColor = GetString(specification, "HeaderTextColor"),
            TextColor = GetString(specification, "TextColor"),
            RowStripeFill = GetString(specification, "RowStripeFill"),
            BorderColor = GetString(specification, "BorderColor"),
            BorderWidth = GetDouble(specification, "BorderWidth"),
            FontSize = GetDouble(specification, "FontSize"),
            HeaderFontSize = GetDouble(specification, "HeaderFontSize"),
            LineHeight = GetDouble(specification, "LineHeight"),
            CellPaddingX = GetDouble(specification, "CellPaddingX", "PaddingX"),
            CellPaddingY = GetDouble(specification, "CellPaddingY", "PaddingY"),
            SpacingBefore = GetDouble(specification, "SpacingBefore"),
            SpacingAfter = GetDouble(specification, "SpacingAfter"),
            Caption = GetString(specification, "Caption"),
            CaptionAlign = GetEnum<PdfAlign>(specification, "CaptionAlign"),
            CaptionColor = GetString(specification, "CaptionColor"),
            CaptionFontSize = GetDouble(specification, "CaptionFontSize"),
            MinimumShrinkFontSize = GetDouble(specification, "MinimumShrinkFontSize", "MinShrinkFontSize"),
            ColumnWidthPoints = GetDoubleArray(specification, "ColumnWidthPoints", "ColumnWidths"),
            ColumnWidthWeights = GetDoubleArray(specification, "ColumnWidthWeights", "ColumnWeights"),
            ColumnAlign = GetEnumArray<PdfColumnAlign>(specification, "ColumnAlign", "ColumnAlignment"),
            AutoFitColumns = GetBool(specification, "AutoFitColumns"),
            RightAlignNumeric = GetBool(specification, "RightAlignNumeric"),
            ShrinkTextToFit = GetBool(specification, "ShrinkTextToFit", "FitTextToCell"),
            KeepTogether = GetBool(specification, "KeepTogether"),
            KeepWithNext = GetBool(specification, "KeepWithNext"),
            NoBorder = GetBool(specification, "NoBorder"),
            NoHeaderFill = GetBool(specification, "NoHeaderFill"),
            NoRowStripeFill = GetBool(specification, "NoRowStripeFill"),
            HeaderRowCount = GetInt(specification, "HeaderRowCount"),
            RepeatHeaderRowCount = GetInt(specification, "RepeatHeaderRowCount"),
            FooterRowCount = GetInt(specification, "FooterRowCount")
        });
    }

    private static PdfTableStyle ResolvePreset(string? styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
        {
            return new PdfTableStyle();
        }

        return Normalize(styleName!) switch
        {
            "light" => TableStyles.Light(),
            "minimal" => TableStyles.Minimal(),
            "rightalignednumbers" => TableStyles.RightAlignedNumbers(),
            "technicaldocument" or "technical" => TableStyles.TechnicalDocument(),
            "compact" => TableStyles.Compact(),
            "report" => TableStyles.Report(),
            _ => TableStyles.FromWordTableStyle(styleName!)
        };
    }

    private static string? GetString(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static double? GetDouble(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value == null ? null : Convert.ToDouble(value, CultureInfo.InvariantCulture);
    }

    private static int? GetInt(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value == null ? null : Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    private static bool GetBool(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value != null && Convert.ToBoolean(value, CultureInfo.InvariantCulture);
    }

    private static TEnum? GetEnum<TEnum>(object specification, params string[] names)
        where TEnum : struct
    {
        var value = GetValue(specification, names);
        if (value == null)
        {
            return null;
        }

        return value is TEnum typed
            ? typed
            : (TEnum)Enum.Parse(typeof(TEnum), Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, ignoreCase: true);
    }

    private static double[]? GetDoubleArray(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        if (value == null)
        {
            return null;
        }

        if (value is string)
        {
            return new[] { Convert.ToDouble(value, CultureInfo.InvariantCulture) };
        }

        if (value is IEnumerable enumerable)
        {
            return enumerable.Cast<object?>()
                .Where(item => item != null)
                .Select(item => Convert.ToDouble(item, CultureInfo.InvariantCulture))
                .ToArray();
        }

        return new[] { Convert.ToDouble(value, CultureInfo.InvariantCulture) };
    }

    private static TEnum[]? GetEnumArray<TEnum>(object specification, params string[] names)
        where TEnum : struct
    {
        var value = GetValue(specification, names);
        if (value == null)
        {
            return null;
        }

        if (value is string || value is TEnum)
        {
            return new[] { ConvertEnum<TEnum>(value) };
        }

        if (value is IEnumerable enumerable)
        {
            return enumerable.Cast<object?>()
                .Where(item => item != null)
                .Select(item => ConvertEnum<TEnum>(item!))
                .ToArray();
        }

        return new[] { ConvertEnum<TEnum>(value) };
    }

    private static TEnum ConvertEnum<TEnum>(object value)
        where TEnum : struct
    {
        return value is TEnum typed
            ? typed
            : (TEnum)Enum.Parse(typeof(TEnum), Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, ignoreCase: true);
    }

    private static object? GetValue(object specification, params string[] names)
    {
        if (specification is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (names.Any(name => string.Equals(name, key, StringComparison.OrdinalIgnoreCase)))
                {
                    return entry.Value;
                }
            }
        }

        var psObject = PSObject.AsPSObject(specification);
        foreach (var name in names)
        {
            var property = psObject.Properties
                .Cast<PSPropertyInfo>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase));
            if (property != null)
            {
                return property.Value;
            }
        }

        return null;
    }

    private static string Normalize(string value)
    {
        return value.Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();
    }
}
