using System;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>PowerShell-friendly description of a PowerPoint shape.</summary>
public sealed class PowerPointShapeInfo
{
    /// <summary>The slide the shape belongs to.</summary>
    public PowerPointSlide Slide { get; set; } = null!;

    /// <summary>The underlying OfficeIMO shape wrapper.</summary>
    public PowerPointShape Shape { get; set; } = null!;

    /// <summary>Zero-based slide index when known.</summary>
    public int SlideIndex { get; set; } = -1;

    /// <summary>Zero-based shape index on the slide.</summary>
    public int ShapeIndex { get; set; } = -1;

    /// <summary>High-level PowerPoint shape kind.</summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>Shape name.</summary>
    public string? Name { get; set; }

    /// <summary>Shape text for textbox shapes.</summary>
    public string? Text { get; set; }

    /// <summary>Paragraph count for textbox shapes.</summary>
    public int ParagraphCount { get; set; }

    /// <summary>Whether the textbox is a slide placeholder.</summary>
    public bool IsPlaceholder { get; set; }

    /// <summary>Placeholder type for placeholder textboxes.</summary>
    public string? PlaceholderType { get; set; }

    /// <summary>Placeholder index for placeholder textboxes.</summary>
    public uint? PlaceholderIndex { get; set; }

    /// <summary>Table row count.</summary>
    public int RowCount { get; set; }

    /// <summary>Table column count.</summary>
    public int ColumnCount { get; set; }

    /// <summary>Image MIME type for picture shapes.</summary>
    public string? MimeType { get; set; }

    /// <summary>Auto shape preset type when available.</summary>
    public string? AutoShapeType { get; set; }

    /// <summary>Left position in points.</summary>
    public double LeftPoints { get; set; }

    /// <summary>Top position in points.</summary>
    public double TopPoints { get; set; }

    /// <summary>Width in points.</summary>
    public double WidthPoints { get; set; }

    /// <summary>Height in points.</summary>
    public double HeightPoints { get; set; }
}

internal static class PowerPointShapeReader
{
    public static PowerPointShapeInfo Read(PowerPointSlide slide, PowerPointShape shape, int slideIndex, int shapeIndex)
    {
        if (slide == null)
        {
            throw new ArgumentNullException(nameof(slide));
        }

        if (shape == null)
        {
            throw new ArgumentNullException(nameof(shape));
        }

        var info = new PowerPointShapeInfo
        {
            Slide = slide,
            Shape = shape,
            SlideIndex = slideIndex,
            ShapeIndex = shapeIndex,
            Kind = GetKind(shape),
            Name = shape.Name,
            LeftPoints = shape.LeftPoints,
            TopPoints = shape.TopPoints,
            WidthPoints = shape.WidthPoints,
            HeightPoints = shape.HeightPoints
        };

        switch (shape)
        {
            case PowerPointTextBox textBox:
                info.Text = textBox.Text;
                info.ParagraphCount = textBox.Paragraphs.Count;
                info.IsPlaceholder = textBox.IsPlaceholder;
                info.PlaceholderType = textBox.PlaceholderType?.ToString();
                info.PlaceholderIndex = textBox.PlaceholderIndex;
                break;
            case PowerPointTable table:
                info.RowCount = table.Rows;
                info.ColumnCount = table.Columns;
                break;
            case PowerPointPicture picture:
                info.MimeType = picture.MimeType;
                break;
            case PowerPointAutoShape autoShape:
                info.AutoShapeType = autoShape.ShapeType?.ToString();
                break;
        }

        return info;
    }

    public static string GetKind(PowerPointShape shape)
    {
        return shape switch
        {
            PowerPointTextBox => "TextBox",
            PowerPointPicture => "Picture",
            PowerPointTable => "Table",
            PowerPointChart => "Chart",
            PowerPointAutoShape => "AutoShape",
            PowerPointGroupShape => "GroupShape",
            _ => shape.GetType().Name
        };
    }
}
