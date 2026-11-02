using System;
using System.Linq;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>PowerShell-friendly summary of a PowerPoint slide.</summary>
public sealed class PowerPointSlideSummaryInfo
{
    /// <summary>The slide the summary belongs to.</summary>
    public PowerPointSlide? Slide { get; set; }

    /// <summary>The inspected source file; null for a live document.</summary>
    public string? Path { get; set; }

    /// <summary>One-based slide position when known.</summary>
    public int SlideIndex { get; set; } = -1;

    /// <summary>One-based master position when known.</summary>
    public int MasterIndex { get; set; } = -1;

    /// <summary>One-based layout position when known.</summary>
    public int LayoutIndex { get; set; } = -1;

    /// <summary>Slide layout name when known.</summary>
    public string? LayoutName { get; set; }

    /// <summary>Slide layout type when known.</summary>
    public string? LayoutType { get; set; }

    /// <summary>Resolved title with physical line breaks normalized to spaces.</summary>
    public string? Title { get; set; }

    /// <summary>Whether the slide already has speaker notes.</summary>
    public bool HasNotes { get; set; }

    /// <summary>Speaker notes text.</summary>
    public string NotesText { get; set; } = string.Empty;

    /// <summary>Total shape count.</summary>
    public int ShapeCount { get; set; }

    /// <summary>Textbox count.</summary>
    public int TextBoxCount { get; set; }

    /// <summary>Picture count.</summary>
    public int PictureCount { get; set; }

    /// <summary>Table count.</summary>
    public int TableCount { get; set; }

    /// <summary>Chart count.</summary>
    public int ChartCount { get; set; }

    /// <summary>Placeholder textbox count.</summary>
    public int PlaceholderCount { get; set; }

    /// <summary>Layout placeholder count.</summary>
    public int LayoutPlaceholderCount { get; set; }
}

internal static class PowerPointSlideSummaryReader
{
    public static PowerPointSlideSummaryInfo Read(PowerPointSlide slide, int slideIndex, string? path = null)
    {
        if (slide == null)
        {
            throw new ArgumentNullException(nameof(slide));
        }

        var notes = PowerPointNotesReader.Read(slide, slideIndex);
        var layout = slide.GetLayoutInfo();

        return new PowerPointSlideSummaryInfo
        {
            Slide = path == null ? slide : null,
            Path = path,
            SlideIndex = slideIndex + 1,
            MasterIndex = layout.HasValue ? layout.Value.MasterIndex + 1 : -1,
            LayoutIndex = slide.LayoutIndex + 1,
            LayoutName = layout?.Name,
            LayoutType = layout?.Type?.ToString(),
            Title = TitleLine(ResolveTitle(slide)),
            HasNotes = notes.HasNotes,
            NotesText = notes.Text,
            ShapeCount = slide.Shapes.Count,
            TextBoxCount = slide.TextBoxes.Count(),
            PictureCount = slide.Pictures.Count(),
            TableCount = slide.Tables.Count(),
            ChartCount = slide.Charts.Count(),
            PlaceholderCount = slide.TextBoxes.Count(textBox => textBox.IsPlaceholder),
            LayoutPlaceholderCount = slide.GetLayoutPlaceholders().Count
        };
    }

    private static string? ResolveTitle(PowerPointSlide slide)
    {
        var title = slide.GetPlaceholder(PowerPointPlaceholderType.Title)?.Text;
        if (!string.IsNullOrWhiteSpace(title))
        {
            return title;
        }

        title = slide.GetPlaceholder(PowerPointPlaceholderType.CenteredTitle)?.Text;
        if (!string.IsNullOrWhiteSpace(title))
        {
            return title;
        }

        return slide.TextBoxes
            .Select(textBox => textBox.Text)
            .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text));
    }
    private static string? TitleLine(string? text) => text == null ? null : string.Join(" ",
        text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Select(line => line.Trim()).Where(line => line.Length > 0));

}
