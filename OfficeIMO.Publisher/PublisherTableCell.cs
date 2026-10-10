using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher;

/// <summary>A native table cell, including its merged span and recovered styled paragraphs.</summary>
public sealed class PublisherTableCell {
    internal PublisherTableCell(int row, int column, int rowSpan, int columnSpan,
        double x, double y, double width, double height, IReadOnlyList<OfficeRichTextParagraph> paragraphs) {
        RowIndex = row; ColumnIndex = column; RowSpan = rowSpan; ColumnSpan = columnSpan;
        X = x; Y = y; Width = width; Height = height;
        Paragraphs = Array.AsReadOnly(paragraphs.ToArray());
        Text = string.Join("\n", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));
    }
    /// <summary>Zero-based first row occupied by this cell.</summary>
    public int RowIndex { get; }
    /// <summary>Zero-based first column occupied by this cell.</summary>
    public int ColumnIndex { get; }
    /// <summary>Number of grid rows occupied by this cell.</summary>
    public int RowSpan { get; }
    /// <summary>Number of grid columns occupied by this cell.</summary>
    public int ColumnSpan { get; }
    /// <summary>Left edge in table-local points.</summary>
    public double X { get; }
    /// <summary>Top edge in table-local points.</summary>
    public double Y { get; }
    /// <summary>Width of the occupied column tracks in points.</summary>
    public double Width { get; }
    /// <summary>Height of the occupied row tracks in points.</summary>
    public double Height { get; }
    /// <summary>Recovered paragraphs and styled runs. Empty when the owning table has no resolved text mapping.</summary>
    public IReadOnlyList<OfficeRichTextParagraph> Paragraphs { get; }
    /// <summary>Recovered cell text with paragraphs separated by newlines. Empty when text mapping is unresolved.</summary>
    public string Text { get; }
}
