namespace OfficeIMO.Reader.Subtitles;

/// <summary>Controls bounded SRT and WebVTT projection.</summary>
public sealed class ReaderSubtitleOptions {
    /// <summary>Gets or sets whether cue timing is included in Markdown. Default: true.</summary>
    public bool IncludeTimestampsInMarkdown { get; set; } = true;

    /// <summary>Gets or sets whether HTML/WebVTT cue markup is removed from text. Default: true.</summary>
    public bool StripCueMarkup { get; set; } = true;

    /// <summary>Gets or sets the maximum number of cues emitted. Default: 50,000.</summary>
    public int MaxCues { get; set; } = 50_000;

    /// <summary>Gets or sets the maximum characters retained for one cue. Default: 32,000.</summary>
    public int MaxCueCharacters { get; set; } = 32_000;

    internal ReaderSubtitleOptions CloneValidated() {
        if (MaxCues < 1 || MaxCues > 1_000_000) throw new ArgumentOutOfRangeException(nameof(MaxCues));
        if (MaxCueCharacters < 1 || MaxCueCharacters > 1_000_000) throw new ArgumentOutOfRangeException(nameof(MaxCueCharacters));
        return new ReaderSubtitleOptions {
            IncludeTimestampsInMarkdown = IncludeTimestampsInMarkdown,
            StripCueMarkup = StripCueMarkup,
            MaxCues = MaxCues,
            MaxCueCharacters = MaxCueCharacters
        };
    }
}
