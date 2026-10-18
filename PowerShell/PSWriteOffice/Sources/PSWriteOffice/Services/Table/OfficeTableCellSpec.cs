using System;
using System.Collections.Generic;
using System.Linq;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.Table;

/// <summary>Describes a logical table cell that can be rendered by multiple Office table surfaces.</summary>
public sealed class OfficeTableCellSpec
{
    /// <summary>Creates a logical table cell.</summary>
    /// <param name="text">Cell text.</param>
    /// <param name="columnSpan">Number of logical columns covered by the cell.</param>
    /// <param name="rowSpan">Number of logical rows covered by the cell.</param>
    /// <param name="style">Optional cell-level style hints consumed by supported table renderers.</param>
    /// <param name="runs">Optional rich text runs used as the cell text when supported by the renderer.</param>
    public OfficeTableCellSpec(string? text, int columnSpan = 1, int rowSpan = 1, OfficeTableCellStyle? style = null, IReadOnlyList<OfficeTextRunSpec>? runs = null)
        : this(text, columnSpan, rowSpan, style, runs, nativeCell: null)
    {
    }

    internal OfficeTableCellSpec(
        string? text,
        int columnSpan,
        int rowSpan,
        OfficeTableCellStyle? style,
        IReadOnlyList<OfficeTextRunSpec>? runs,
        object? nativeCell)
    {
        if (columnSpan < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columnSpan), "Column span must be at least 1.");
        }

        if (rowSpan < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rowSpan), "Row span must be at least 1.");
        }

        Runs = runs?.ToArray();
        Text = text ?? (Runs is { Count: > 0 } ? OfficeTextRunParser.GetPlainText(Runs.ToArray()) : string.Empty);
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        Style = style;
        NativeCell = nativeCell;
    }

    /// <summary>Cell text.</summary>
    public string Text { get; }

    /// <summary>Number of logical columns covered by the cell.</summary>
    public int ColumnSpan { get; }

    /// <summary>Number of logical rows covered by the cell.</summary>
    public int RowSpan { get; }

    /// <summary>Optional cell-level style hints.</summary>
    public OfficeTableCellStyle? Style { get; }

    /// <summary>Optional rich text runs for renderers that support inline cell formatting.</summary>
    public IReadOnlyList<OfficeTextRunSpec>? Runs { get; }

    /// <summary>Renderer-owned typed cell payload retained without widening the cross-format public contract.</summary>
    internal object? NativeCell { get; }

    internal bool HasSpan => ColumnSpan > 1 || RowSpan > 1;

    internal bool HasStyle => Style?.HasAnyValue == true;

    internal bool HasRuns => Runs is { Count: > 0 };

    internal bool HasNativeCell => NativeCell != null;

    internal bool HasStructuredMarker => HasSpan || HasStyle || HasRuns || HasNativeCell;
}

/// <summary>Optional logical table cell style hints for renderers that support per-cell formatting.</summary>
public sealed class OfficeTableCellStyle
{
    /// <summary>Cell text color. Named colors and hexadecimal colors are accepted.</summary>
    public string? TextColor { get; set; }

    /// <summary>Cell fill color. Named colors and hexadecimal colors are accepted.</summary>
    public string? FillColor { get; set; }

    /// <summary>Cell font size in points.</summary>
    public double? FontSize { get; set; }

    /// <summary>Render cell text in bold.</summary>
    public bool Bold { get; set; }

    /// <summary>Render cell text in italics.</summary>
    public bool Italic { get; set; }

    /// <summary>Render cell text with underline.</summary>
    public bool Underline { get; set; }

    /// <summary>Optional underline style name when the target renderer supports it.</summary>
    public string? UnderlineStyle { get; set; }

    /// <summary>Render cell text with strikethrough.</summary>
    public bool Strike { get; set; }

    /// <summary>Horizontal alignment name used by supported renderers.</summary>
    public string? Align { get; set; }

    /// <summary>Vertical alignment name used by supported renderers.</summary>
    public string? VerticalAlign { get; set; }

    internal bool HasAnyValue =>
        !string.IsNullOrWhiteSpace(TextColor) ||
        !string.IsNullOrWhiteSpace(FillColor) ||
        FontSize.HasValue ||
        Bold ||
        Italic ||
        Underline ||
        !string.IsNullOrWhiteSpace(UnderlineStyle) ||
        Strike ||
        !string.IsNullOrWhiteSpace(Align) ||
        !string.IsNullOrWhiteSpace(VerticalAlign);

    internal bool HasTextStyle =>
        !string.IsNullOrWhiteSpace(TextColor) ||
        FontSize.HasValue ||
        Bold ||
        Italic ||
        Underline ||
        !string.IsNullOrWhiteSpace(UnderlineStyle) ||
        Strike;

    internal bool HasTableStyle =>
        !string.IsNullOrWhiteSpace(FillColor) ||
        !string.IsNullOrWhiteSpace(Align) ||
        !string.IsNullOrWhiteSpace(VerticalAlign);
}
