using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>PowerShell-friendly description of slide speaker notes.</summary>
public sealed class PowerPointNotesInfo
{
    /// <summary>The slide the notes belong to.</summary>
    public PowerPointSlide Slide { get; set; } = null!;

    /// <summary>Zero-based slide index when known.</summary>
    public int SlideIndex { get; set; } = -1;

    /// <summary>Whether a notes part already exists on the slide.</summary>
    public bool HasNotes { get; set; }

    /// <summary>Speaker notes text.</summary>
    public string Text { get; set; } = string.Empty;
}

internal static class PowerPointNotesReader
{
    private static readonly PropertyInfo? SlidePartProperty = typeof(PowerPointSlide).GetProperty("SlidePart", BindingFlags.Instance | BindingFlags.NonPublic);

    public static PowerPointNotesInfo Read(PowerPointSlide slide, int slideIndex)
    {
        if (slide == null)
        {
            throw new ArgumentNullException(nameof(slide));
        }

        var text = ReadTextNoCreate(slide, out var hasNotes);
        return new PowerPointNotesInfo
        {
            Slide = slide,
            SlideIndex = slideIndex,
            HasNotes = hasNotes,
            Text = text
        };
    }

    public static int ResolveSlideIndex(PowerPointSlide slide)
    {
        if (slide == null)
        {
            throw new ArgumentNullException(nameof(slide));
        }

        var slidePart = ResolveSlidePart(slide);
        var presentationPart = slidePart?.GetParentParts().OfType<PresentationPart>().FirstOrDefault();
        var slideIds = presentationPart?.Presentation?.SlideIdList?.Elements<SlideId>();
        if (slidePart == null || slideIds == null)
        {
            return -1;
        }

        var index = 0;
        foreach (var slideId in slideIds)
        {
            var relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                index++;
                continue;
            }

            if (ReferenceEquals(presentationPart!.GetPartById(relationshipId!), slidePart))
            {
                return index;
            }

            index++;
        }

        return -1;
    }

    private static string ReadTextNoCreate(PowerPointSlide slide, out bool hasNotes)
    {
        hasNotes = false;

        try
        {
            var slidePart = ResolveSlidePart(slide);
            var notesSlide = slidePart?.NotesSlidePart?.NotesSlide;
            if (notesSlide == null)
            {
                return string.Empty;
            }

            hasNotes = true;
            var blocks = notesSlide.CommonSlideData?.ShapeTree?
                .Elements<Shape>()
                .Select(ReadShapeText)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList() ?? new List<string>();

            return string.Join(Environment.NewLine + Environment.NewLine, blocks);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string ReadShapeText(Shape shape)
    {
        var paragraphs = shape.TextBody?
            .Elements<A.Paragraph>()
            .Select(ReadParagraphText)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList() ?? new List<string>();

        return string.Join(Environment.NewLine, paragraphs);
    }

    private static string ReadParagraphText(A.Paragraph paragraph)
    {
        var builder = new StringBuilder();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case A.Run run:
                    builder.Append(run.Text?.Text ?? string.Empty);
                    break;
                case A.Break:
                    builder.AppendLine();
                    break;
                case A.Field field:
                    builder.Append(field.Text?.Text ?? string.Empty);
                    break;
            }
        }

        return (builder.ToString() ?? string.Empty).Trim();
    }

    internal static SlidePart? ResolveSlidePart(PowerPointSlide slide)
    {
        return SlidePartProperty?.GetValue(slide) as SlidePart;
    }
}
