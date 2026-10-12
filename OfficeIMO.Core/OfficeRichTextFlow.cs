using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Continues immutable source paragraphs through successive measured regions.
/// This adapter primitive uses the same line breaker as drawing renderers and
/// retains source whitespace and UTF-16 boundaries instead of parsing painted text.
/// </summary>
internal sealed class OfficeRichTextFlow {
    private readonly IReadOnlyList<OfficeRichTextParagraph> _paragraphs;
    private readonly Action<int> _accountMeasurement;
    private int _paragraph, _offset;

    internal OfficeRichTextFlow(IReadOnlyList<OfficeRichTextParagraph> paragraphs, Action<int> accountMeasurement) {
        _paragraphs = paragraphs ?? throw new ArgumentNullException(nameof(paragraphs));
        _accountMeasurement = accountMeasurement ?? throw new ArgumentNullException(nameof(accountMeasurement));
    }

    /// <summary>Position in the source paragraphs joined by a single newline, before tab expansion.</summary>
    internal int CharacterPosition { get; private set; }
    internal bool HasRemaining => _paragraph < _paragraphs.Count;

    /// <summary>
    /// Takes complete measured lines that fit the region. A continuation does
    /// not repeat first-line indentation, a list label, or space before its paragraph.
    /// Space after a completed paragraph is dropped at a region boundary.
    /// </summary>
    internal IReadOnlyList<OfficeRichTextParagraph> Take(double width, double height,
        Func<string?, double, string?, OfficeFontStyle, double> measure, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var visible = new List<OfficeRichTextParagraph>();
        if (width <= 0 || height <= 0) return visible;
        double remainingHeight = height;
        while (HasRemaining && remainingHeight > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRichTextParagraph source = _paragraphs[_paragraph];
            int length = TextLength(source);
            OfficeRichTextParagraph tail = Slice(source, _offset, length - _offset, complete: true);
            _accountMeasurement(length - _offset);
            OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { tail },
                width, remainingHeight, measure, cancellationToken: cancellationToken);
            int consumed = 0;
            bool sourceEndVisible = false;
            foreach (OfficeRichTextLine line in layout.Lines) {
                // Ordinary drawing wrapping may force an indivisible glyph onto
                // a narrow line. Flow must retain it for a region that can hold it.
                if (line.OffsetX < -.000001D || line.OffsetX + line.Width > width - tail.Margins.Right + .000001D) break;
                consumed = Math.Max(consumed, line.SourceTextEnd);
                sourceEndVisible |= line.CompletesSource;
            }
            consumed = Math.Min(consumed, length - _offset);
            bool complete = consumed == length - _offset && sourceEndVisible;
            // An empty paragraph still consumes a printable line; outer margin
            // gaps alone do not establish that the empty line fitted.
            if (consumed == 0 && !complete) break;
            visible.Add(Slice(source, _offset, consumed, complete));
            CharacterPosition = checked(CharacterPosition + consumed);
            _offset += consumed;
            remainingHeight -= layout.Height;
            if (!complete) break;
            _paragraph++; _offset = 0;
            if (HasRemaining) CharacterPosition++;
        }
        return visible;
    }

    private static int TextLength(OfficeRichTextParagraph paragraph) {
        int length = 0;
        foreach (OfficeRichTextRun run in paragraph.Runs) length = checked(length + run.Text.Length);
        return length;
    }

    private static OfficeRichTextParagraph Slice(OfficeRichTextParagraph source, int start, int length, bool complete) {
        var runs = new List<OfficeRichTextRun>();
        int position = 0, end = checked(start + length);
        foreach (OfficeRichTextRun run in source.Runs) {
            int next = checked(position + run.Text.Length);
            int left = Math.Max(position, start), right = Math.Min(next, end);
            if (right > left) {
                OfficeTextParagraphIndent? indent = run.ParagraphIndent;
                if (start != 0 && indent.HasValue) indent = new OfficeTextParagraphIndent(
                    indent.Value.ContinuationLineOffset, indent.Value.ContinuationLineOffset);
                runs.Add(new OfficeRichTextRun(run.Text.Substring(left - position, right - left),
                    run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily,
                    run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle,
                    run.Baseline, indent) { LinkUri = run.LinkUri });
            }
            position = next;
        }
        if (runs.Count == 0 && source.Runs.Count > 0) {
            // Shared layout measures an empty source line with the paragraph's
            // maximum effective run size. Keep its style set when only the
            // trailing blank line remains, including already consumed runs.
            foreach (OfficeRichTextRun run in source.Runs)
                runs.Add(new OfficeRichTextRun(string.Empty, run.FontSize, run.Color, run.Bold, run.Italic,
                    run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor,
                    run.UnderlineStyle, run.StrikethroughStyle, run.Baseline));
        }
        var margins = new OfficeTextPadding(source.Margins.Left, start == 0 ? source.Margins.Top : 0,
            source.Margins.Right, complete ? source.Margins.Bottom : 0);
        OfficeTextParagraphIndent paragraphIndent = start == 0 ? source.Indent :
            new OfficeTextParagraphIndent(source.Indent.ContinuationLineOffset, source.Indent.ContinuationLineOffset);
        OfficeRichTextParagraph result = start == 0 && source.Label != null
            ? new OfficeRichTextParagraph(runs, source.Label, source.Alignment, source.LineHeight, margins, paragraphIndent, source.LineHeightFactor)
            : new OfficeRichTextParagraph(runs, source.Alignment, source.LineHeight, margins, paragraphIndent, source.LineHeightFactor);
        result.ContinuesInNextRegion = !complete || source.ContinuesInNextRegion;
        return result.WithTabStops(source.TabStops);
    }
}
