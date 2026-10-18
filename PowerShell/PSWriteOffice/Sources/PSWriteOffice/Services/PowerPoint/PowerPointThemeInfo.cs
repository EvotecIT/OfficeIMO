using System.Collections.Generic;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Describes a PowerPoint theme for a single slide master.</summary>
public sealed class PowerPointThemeInfo
{
    /// <summary>Creates a new theme info snapshot.</summary>
    public PowerPointThemeInfo(
        int masterIndex,
        string themeName,
        IReadOnlyDictionary<PowerPointThemeColor, string> colors,
        PowerPointThemeFontSet fonts)
    {
        MasterIndex = masterIndex;
        ThemeName = themeName ?? string.Empty;
        Colors = colors;
        Fonts = fonts;
    }

    /// <summary>Slide master index that was inspected.</summary>
    public int MasterIndex { get; }

    /// <summary>Theme name associated with the presentation.</summary>
    public string ThemeName { get; }

    /// <summary>Theme colors keyed by <see cref="PowerPointThemeColor"/>.</summary>
    public IReadOnlyDictionary<PowerPointThemeColor, string> Colors { get; }

    /// <summary>Theme font set for the selected master.</summary>
    public PowerPointThemeFontSet Fonts { get; }

    /// <summary>Major Latin theme font.</summary>
    public string? MajorLatin => Fonts.MajorLatin;

    /// <summary>Minor Latin theme font.</summary>
    public string? MinorLatin => Fonts.MinorLatin;

    /// <summary>Major East Asian theme font.</summary>
    public string? MajorEastAsian => Fonts.MajorEastAsian;

    /// <summary>Minor East Asian theme font.</summary>
    public string? MinorEastAsian => Fonts.MinorEastAsian;

    /// <summary>Major complex-script theme font.</summary>
    public string? MajorComplexScript => Fonts.MajorComplexScript;

    /// <summary>Minor complex-script theme font.</summary>
    public string? MinorComplexScript => Fonts.MinorComplexScript;
}
