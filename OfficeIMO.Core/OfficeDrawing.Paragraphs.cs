using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>Adds paragraphs with independent styles and render-time font measurement inside one text frame.</summary>
    /// <remarks>Input is limited to 100,000 characters and 4,096 runs including paragraph separators. Layout omits paragraphs whose horizontal margins consume the frame and lines beyond the frame height; glyph ink may extend beyond condensed line spacing. Native auto-sizing is not inferred.</remarks>
    public OfficeDrawing AddRichTextParagraphs(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double x, double y, double width, double height, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        bool wrapText = true, OfficeTextPadding? padding = null) =>
        AddRichTextParagraphsCore(paragraphs, x, y, width, height, verticalAlignment, wrapText, padding, shrinkToFit: false);

    /// <summary>Adds independently aligned paragraphs in an intrinsically sized text area inside one drawing frame.</summary>
    /// <remarks>The measured area includes paragraph margins and indentation and excludes frame padding. Wrapping uses the padded frame width before placement; shorter lines retain their paragraph alignment inside the measured area. Height clipping does not change the area width, and an unwrapped area wider than the frame retains its requested anchor.</remarks>
    public OfficeDrawing AddRichTextParagraphs(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double x, double y, double width, double height, OfficeTextAreaAlignment areaAlignment,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        bool wrapText = true, OfficeTextPadding? padding = null) =>
        AddRichTextParagraphsCore(paragraphs, x, y, width, height, verticalAlignment, wrapText, padding,
            shrinkToFit: false, areaAlignment: areaAlignment);

    // Format adapters can request the existing shared fitting contract without changing the public default.
    internal OfficeDrawing AddRichTextParagraphsCore(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double x, double y, double width, double height, OfficeTextVerticalAlignment verticalAlignment,
        bool wrapText, OfficeTextPadding? padding, bool shrinkToFit) =>
        AddRichTextParagraphsCore(paragraphs, x, y, width, height, verticalAlignment, wrapText, padding,
            shrinkToFit, OfficeTextAreaAlignment.FullWidth);

    // Preserve the original CLR signature used by independently packaged friend assemblies.
    internal OfficeDrawing AddRichTextParagraphsCore(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double x, double y, double width, double height, OfficeTextVerticalAlignment verticalAlignment,
        bool wrapText, OfficeTextPadding? padding, bool shrinkToFit,
        OfficeTextAreaAlignment areaAlignment) {
        if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));
        if (!Enum.IsDefined(typeof(OfficeTextAreaAlignment), areaAlignment)) throw new ArgumentOutOfRangeException(nameof(areaAlignment));
        if (paragraphs.Count > OfficeTextLayoutEngine.MaximumLayoutLines) throw new ArgumentException("Paragraph count exceeds the shared line limit.", nameof(paragraphs));
        var runs = new List<OfficeRichTextRun>();
        int characters = 0;
        for (int i = 0; i < paragraphs.Count; i++) {
            OfficeRichTextParagraph paragraph = paragraphs[i] ?? throw new ArgumentException("Paragraphs cannot contain null entries.", nameof(paragraphs));
            if (i > 0) { runs.Add(new OfficeRichTextRun("\n", 1, OfficeColor.Black)); characters++; }
            if (paragraph.Label != null) {
                var label = paragraph.Label;
                string value = label.Run.Text + (label.FollowedBy == OfficeTextParagraphLabelFollowedBy.Nothing ? "" : " ");
                if (runs.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns || value.Length > OfficeTextLayoutEngine.MaximumLayoutTextCharacters - characters)
                    throw new ArgumentException("Paragraph labels exceed the shared rich text limit.", nameof(paragraphs));
                characters += value.Length;
                var copy = OfficeRichTextParagraph.CopyRun(label.Run);
                runs.Add(new OfficeRichTextRun(value, copy.FontSize, copy.Color, copy.Bold, copy.Italic, copy.Underline, copy.FontFamily,
                    copy.Strikethrough, copy.BackgroundColor, copy.UnderlineStyle, copy.StrikethroughStyle, copy.Baseline) { LinkUri = copy.LinkUri });
            }
            foreach (OfficeRichTextRun run in paragraph.Runs) {
                if (runs.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns || run.Text.Length > OfficeTextLayoutEngine.MaximumLayoutTextCharacters - characters)
                    throw new ArgumentException("Paragraph text exceeds the shared rich text run or character limit.", nameof(paragraphs));
                characters = checked(characters + run.Text.Length);
                runs.Add(run);
            }
            if (runs.Count > OfficeTextLayoutEngine.MaximumLayoutTextRuns || characters > OfficeTextLayoutEngine.MaximumLayoutTextCharacters)
                throw new ArgumentException("Paragraph text exceeds the shared rich text run or character limit.", nameof(paragraphs));
        }
        var item = new OfficeDrawingRichText(runs, x, y, width, height,
            verticalAlignment: verticalAlignment, wrapText: wrapText, padding: padding, shrinkToFit: shrinkToFit)
            .WithParagraphs(paragraphs).WithTextAreaAlignment(areaAlignment);
        if (item.X + item.Width > Width || item.Y + item.Height > Height)
            throw new ArgumentOutOfRangeException(nameof(paragraphs), "Drawing paragraph text must fit inside the drawing bounds.");
        _elements.Add(item); return this;
    }
}
