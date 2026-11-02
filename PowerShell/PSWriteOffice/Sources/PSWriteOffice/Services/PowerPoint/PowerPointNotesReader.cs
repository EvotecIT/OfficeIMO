using System;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>PowerShell-friendly description of slide speaker notes.</summary>
public sealed class PowerPointNotesInfo
{
    /// <summary>The slide the notes belong to.</summary>
    public PowerPointSlide Slide { get; set; } = null!;

    /// <summary>One-based slide position when known.</summary>
    public int SlideIndex { get; set; } = -1;

    /// <summary>Whether a notes part already exists on the slide.</summary>
    public bool HasNotes { get; set; }

    /// <summary>Speaker notes text.</summary>
    public string Text { get; set; } = string.Empty;
}

internal static class PowerPointNotesReader {
    public static PowerPointNotesInfo Read(PowerPointSlide slide, int slideIndex) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        return new PowerPointNotesInfo {
            Slide = slide,
            SlideIndex = slideIndex >= 0 ? slideIndex + 1 : -1,
            HasNotes = slide.HasSpeakerNotes,
            Text = slide.GetSpeakerNotesText()
        };
    }

    public static int ResolveSlideIndex(PowerPointSlide slide) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        return slide.Index;
    }
}
