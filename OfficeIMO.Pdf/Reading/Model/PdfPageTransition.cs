namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight metadata for a PDF page transition dictionary.
/// </summary>
public sealed class PdfPageTransition {
    internal PdfPageTransition(string? style, double? durationSeconds, string? dimension, string? motion, int? direction, double? scale, bool? isFlyAreaOpaque) {
        Style = style;
        DurationSeconds = durationSeconds;
        Dimension = dimension;
        Motion = motion;
        Direction = direction;
        Scale = scale;
        IsFlyAreaOpaque = isFlyAreaOpaque;
    }

    /// <summary>Transition style from /S, for example Split, Blinds, Box, Wipe, Dissolve, Glitter, R, Fly, Push, Cover, Uncover, or Fade.</summary>
    public string? Style { get; }

    /// <summary>Transition duration from /D, in seconds, when present.</summary>
    public double? DurationSeconds { get; }

    /// <summary>Transition dimension from /Dm, usually H or V.</summary>
    public string? Dimension { get; }

    /// <summary>Transition motion from /M, usually I or O.</summary>
    public string? Motion { get; }

    /// <summary>Transition direction from /Di, in degrees or PDF-defined special value, when present.</summary>
    public int? Direction { get; }

    /// <summary>Fly transition scale from /SS, when present.</summary>
    public double? Scale { get; }

    /// <summary>Fly transition area opacity flag from /B, when present.</summary>
    public bool? IsFlyAreaOpaque { get; }
}
