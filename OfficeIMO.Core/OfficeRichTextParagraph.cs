using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>A paragraph whose styled runs and layout are resolved with the drawing's render-time fonts.</summary>
public sealed class OfficeRichTextParagraph {
    /// <summary>Creates a paragraph with independent alignment, line spacing, margins and indentation.</summary>
    /// <param name="runs">Styled inline content. The collection and runs are copied.</param>
    /// <param name="alignment">Horizontal alignment within the paragraph's margins.</param>
    /// <param name="lineHeight">Optional absolute line height; null uses the shared default.</param>
    /// <param name="margins">Nonnegative paragraph insets. Adjacent vertical margins are added.</param>
    /// <param name="indent">Offsets of the first and subsequent visual lines.</param>
    /// <param name="lineHeightFactor">Optional multiplier of each line's text extent; mutually exclusive with absolute line height.</param>
    public OfficeRichTextParagraph(IReadOnlyList<OfficeRichTextRun> runs,
        OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null,
        OfficeTextPadding? margins = null, OfficeTextParagraphIndent? indent = null, double? lineHeightFactor = null) {
        if (runs == null) throw new ArgumentNullException(nameof(runs));
        if (!Enum.IsDefined(typeof(OfficeTextAlignment), alignment)) throw new ArgumentOutOfRangeException(nameof(alignment));
        if (lineHeight.HasValue && (lineHeight.Value <= 0 || double.IsNaN(lineHeight.Value) || double.IsInfinity(lineHeight.Value)))
            throw new ArgumentOutOfRangeException(nameof(lineHeight));
        if (lineHeightFactor.HasValue && (lineHeightFactor.Value <= 0 || double.IsNaN(lineHeightFactor.Value) || double.IsInfinity(lineHeightFactor.Value)))
            throw new ArgumentOutOfRangeException(nameof(lineHeightFactor));
        if (lineHeight.HasValue && lineHeightFactor.HasValue) throw new ArgumentException("Choose absolute or relative line height.", nameof(lineHeightFactor));
        if (runs.Count > OfficeTextLayoutEngine.MaximumLayoutTextRuns) throw new ArgumentException("Paragraph exceeds the shared run limit.", nameof(runs));
        var copied = new List<OfficeRichTextRun>(runs.Count);
        int characters = 0;
        foreach (OfficeRichTextRun run in runs) {
            if (run == null) throw new ArgumentException("Paragraph runs cannot contain null entries.", nameof(runs));
            if (run.Text.Length > OfficeTextLayoutEngine.MaximumLayoutTextCharacters - characters)
                throw new ArgumentException("Paragraph exceeds the shared character limit.", nameof(runs));
            characters += run.Text.Length;
            copied.Add(CopyRun(run));
        }
        Runs = new ReadOnlyCollection<OfficeRichTextRun>(copied);
        Alignment = alignment; LineHeight = lineHeight; LineHeightFactor = lineHeightFactor;
        Margins = margins ?? OfficeTextPadding.Empty;
        Indent = indent ?? OfficeTextParagraphIndent.Empty;
    }

    /// <summary>Creates a paragraph with a separately styled label and independent continuation indentation.</summary>
    public OfficeRichTextParagraph(IReadOnlyList<OfficeRichTextRun> runs, OfficeTextParagraphLabel label,
        OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null,
        OfficeTextPadding? margins = null, OfficeTextParagraphIndent? indent = null, double? lineHeightFactor = null)
        : this(runs, alignment, lineHeight, margins, indent, lineHeightFactor) {
        if (label == null) throw new ArgumentNullException(nameof(label));
        int characters = label.Run.Text.Length;
        foreach (OfficeRichTextRun run in Runs) characters = checked(characters + run.Text.Length);
        if (characters > OfficeTextLayoutEngine.MaximumLayoutTextCharacters || Runs.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns)
            throw new ArgumentException("Labelled paragraph exceeds the shared text limit.", nameof(label));
        Label = label.WithRun(label.Run);
    }

    /// <summary>Snapshot of styled inline content.</summary>
    public IReadOnlyList<OfficeRichTextRun> Runs { get; }
    /// <summary>Horizontal alignment within the paragraph's margins.</summary>
    public OfficeTextAlignment Alignment { get; }
    /// <summary>Absolute line height, or null for the shared default.</summary>
    public double? LineHeight { get; }
    /// <summary>Multiplier of each line's text extent, or null for the shared default.</summary>
    public double? LineHeightFactor { get; }
    /// <summary>Paragraph insets; adjacent vertical margins are added.</summary>
    public OfficeTextPadding Margins { get; }
    /// <summary>First-line and continuation-line offsets.</summary>
    public OfficeTextParagraphIndent Indent { get; }
    /// <summary>Optional list label, painted once on the first visual line.</summary>
    public OfficeTextParagraphLabel? Label { get; }

    /// <summary>Optional measured tab settings. Null retains the legacy space-expansion behavior.</summary>
    public OfficeTextTabStops? TabStops { get; private set; }

    /// <summary>Returns an independent paragraph with these tab settings; null restores legacy tab expansion.</summary>
    public OfficeRichTextParagraph WithTabStops(OfficeTextTabStops? tabStops) {
        var copy = Label == null ? new OfficeRichTextParagraph(Runs, Alignment, LineHeight, Margins, Indent, LineHeightFactor) :
            new OfficeRichTextParagraph(Runs, Label, Alignment, LineHeight, Margins, Indent, LineHeightFactor);
        copy.TabStops = tabStops;
        return copy;
    }

    internal static OfficeRichTextRun CopyRun(OfficeRichTextRun run) => new OfficeRichTextRun(
        run.Text, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily,
        run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle,
        run.Baseline, run.ParagraphIndent) { LinkUri = run.LinkUri };
}
