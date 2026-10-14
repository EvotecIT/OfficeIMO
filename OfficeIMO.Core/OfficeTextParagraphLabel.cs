using System;

namespace OfficeIMO.Drawing;

/// <summary>Controls where paragraph text begins after a separately measured label.</summary>
public enum OfficeTextParagraphLabelFollowedBy {
    /// <summary>Text starts at its declared position, or after the label if the label is wider.</summary>
    Position,
    /// <summary>Text starts after one measured space.</summary>
    Space,
    /// <summary>Text starts immediately after the label.</summary>
    Nothing
}

/// <summary>A styled list label placed independently of the paragraph's wrapped text.</summary>
/// <remarks>Render-time layout omits an entire label that exceeds the content rectangle's right edge or right margin,
/// marks the layout as clipped, and preserves the logical text and body insets.</remarks>
public sealed class OfficeTextParagraphLabel {
    private OfficeTextParagraphLabel(OfficeRichTextRun run, double position, OfficeTextAlignment alignment,
        double? minimumWidth, double minimumDistance, OfficeTextParagraphLabelFollowedBy followedBy, double? textPosition) {
        if (run == null) throw new ArgumentNullException(nameof(run));
        if (run.Text.Length == 0 || run.Text.Length > 4096 || run.Text.IndexOfAny(new[] { '\r', '\n', '\t' }) >= 0)
            throw new ArgumentException("A paragraph label requires single-line text of at most 4,096 characters.", nameof(run));
        Validate(position, nameof(position)); Validate(minimumDistance, nameof(minimumDistance));
        if (minimumWidth.HasValue) Validate(minimumWidth.Value, nameof(minimumWidth));
        if (textPosition.HasValue) Validate(textPosition.Value, nameof(textPosition));
        if (alignment is not (OfficeTextAlignment.Left or OfficeTextAlignment.Center or OfficeTextAlignment.Right)) throw new ArgumentOutOfRangeException(nameof(alignment));
        if (!Enum.IsDefined(typeof(OfficeTextParagraphLabelFollowedBy), followedBy)) throw new ArgumentOutOfRangeException(nameof(followedBy));
        Run = OfficeRichTextParagraph.CopyRun(run); Position = position; Alignment = alignment;
        MinimumWidth = minimumWidth; MinimumDistance = minimumDistance; FollowedBy = followedBy; TextPosition = textPosition;
    }

    /// <summary>Places a label inside a minimum-width box. A wider label expands the box.</summary>
    public static OfficeTextParagraphLabel InBox(OfficeRichTextRun run, double position, double minimumWidth,
        OfficeTextAlignment alignment = OfficeTextAlignment.Left, double minimumDistance = 0) =>
        new OfficeTextParagraphLabel(run, position, alignment, minimumWidth, minimumDistance, OfficeTextParagraphLabelFollowedBy.Position, position + minimumWidth);

    /// <summary>Aligns a label at an absolute position inside the text frame's content rectangle.</summary>
    /// <param name="run">The independently styled label.</param>
    /// <param name="position">Left edge, center or right edge according to alignment.</param>
    /// <param name="alignment">Label alignment at the position.</param>
    /// <param name="followedBy">Whether text uses a declared position, a space or no separator.</param>
    /// <param name="textPosition">Optional first-line text position for Position mode.</param>
    public static OfficeTextParagraphLabel AtPosition(OfficeRichTextRun run, double position,
        OfficeTextAlignment alignment = OfficeTextAlignment.Left, OfficeTextParagraphLabelFollowedBy followedBy = OfficeTextParagraphLabelFollowedBy.Position,
        double? textPosition = null) => new OfficeTextParagraphLabel(run, position, alignment, null, 0, followedBy, textPosition);

    /// <summary>Snapshot of the label's styled text.</summary>
    public OfficeRichTextRun Run { get; }
    /// <summary>Position relative to the frame's content rectangle, before scaling.</summary>
    public double Position { get; }
    /// <summary>Alignment inside the box or at the anchor.</summary>
    public OfficeTextAlignment Alignment { get; }
    /// <summary>Minimum label box width, or null for point alignment.</summary>
    public double? MinimumWidth { get; }
    /// <summary>Minimum distance between the label and the first text line.</summary>
    public double MinimumDistance { get; }
    /// <summary>How the first text line follows the label.</summary>
    public OfficeTextParagraphLabelFollowedBy FollowedBy { get; }
    /// <summary>Optional first-line text position, independent of continuation indentation.</summary>
    public double? TextPosition { get; }
    internal OfficeTextParagraphLabel WithRun(OfficeRichTextRun run) => new OfficeTextParagraphLabel(run, Position, Alignment, MinimumWidth, MinimumDistance, FollowedBy, TextPosition);
    private static void Validate(double value, string name) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name, "Label positions and spacing must be finite and nonnegative.");
    }
}
