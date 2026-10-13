using DocumentFormat.OpenXml;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Keeps one non-intersecting page-anchored running-story box out of paragraph
    /// flow. Complex peers, nested boxes and intersecting wrap remain on the
    /// explicitly approximated mixed-content renderer.
    /// </summary>
    private static bool TryRenderNativeAnchoredRunningTextBoxParagraph(INativePdfFlow pdf, WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, (int Level, string Marker)? marker, NativeNoteNumbering notes,
        WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap, WordParagraph? nextParagraph) {
        if (defaults.RunningContentContext == null || !TryPrepareNativeAnchoredRunningTextBoxParagraph(pdf.PageSize,
                paragraph, runs, marker, notes, options, defaults, fontMap, out var prepared)) return false;
        WordParagraph[] selected = prepared!.Runs;
        string text = string.Concat(selected.Select(run => run.Text));
        if (ShouldSuppressNativeContextualSpacingAfter(paragraph, nextParagraph)) prepared.Style.SpacingAfter = 0D;
        if (ShouldSuppressNativeContextualSpacingBefore(paragraph)) prepared.Style.SpacingBefore = 0D;
        pdf.Paragraph(builder => AddNativeParagraphContent(builder, paragraph, null, selected, true, text,
            Array.Empty<int>(), notes, options, defaults, fontMap, projectedContentOnly: true),
            PdfCore.PdfAlign.Left, style: prepared.Style);
        return true;
    }

    private sealed record NativeAnchoredRunningTextBoxParagraph(WordParagraph[] Runs, PdfCore.PdfParagraphStyle Style);

    private static bool TryPrepareNativeAnchoredRunningTextBoxParagraph(PdfCore.PageSize pageSize, WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, (int Level, string Marker)? marker, NativeNoteNumbering? notes,
        WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap,
        out NativeAnchoredRunningTextBoxParagraph? prepared) {
        prepared = null;
        if (defaults.RunningContentSection == null || marker != null || WordDocumentTraversal.GetListInfo(paragraph).HasValue ||
            WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0 ||
            ResolveNativeParagraphAlign(paragraph, allowJustify: false) != PdfCore.PdfAlign.Left) return false;
        options?.CancellationToken.ThrowIfCancellationRequested();
        List<WordTextBox> boxes = runs.SelectMany(run => run.GetTextBoxes()).ToList();
        if (boxes.Count != 1) return false;
        WordTextBox box = boxes[0];
        W.TextBoxContent? content = box.Content;
        if (content == null || content.ChildElements.Count != 1 || content.FirstChild is not W.Paragraph innerSource ||
            content.Descendants<W.TextBoxContent>().Any() || content.Descendants<W.Drawing>().Any() ||
            content.Descendants<W.Picture>().Any() || content.Descendants<W.EmbeddedObject>().Any() ||
            innerSource.ParagraphProperties?.NumberingProperties != null) return false;
        var inner = new WordParagraph(box.Document, innerSource);
        if (WordDocumentTraversal.GetListInfo(inner).HasValue ||
            WordEquation.GetOccurrences(inner._document, inner._paragraph).Count > 0) return false;
        OpenXmlElement[] innerChildren = GetNativeRuns(inner).SelectMany(run => run.EnumerateEffectiveRunContent())
            .Where(child => child is not W.RunProperties).ToArray();
        if (innerChildren.Any(child => child is not W.Text && child is not W.Break) ||
            innerChildren.OfType<W.Break>().Count() > 1 ||
            innerChildren.OfType<W.Break>().Any(lineBreak => !ReferenceEquals(lineBreak, innerChildren.Last()) ||
                lineBreak.Type?.Value == W.BreakValues.Page || lineBreak.Type?.Value == W.BreakValues.Column)) return false;
        if (!TryGetNativeRunningTextBoxFrame(box, defaults.RunningContentSection, options, pageSize,
                out double x, out double y, out double width, out double height)) return false;
        Wps.TextBodyProperties? body = box.DrawingTextBodyProperties;
        if (body == null || body.Rotation?.Value is > 0 or < 0 ||
            body.Vertical?.Value is { } vertical && vertical != A.TextVerticalValues.Horizontal ||
            body.Anchor?.Value is { } anchor && anchor != A.TextAnchoringTypeValues.Top ||
            body.GetFirstChild<A.NormalAutoFit>() != null) return false;

        var surrounding = new List<WordParagraph>();
        foreach (WordParagraph run in runs) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
                if (child is W.RunProperties) continue;
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                if (view.GetTextBoxes().Any()) continue;
                if (child is not W.Text) return false;
                surrounding.Add(view);
            }
        }
        if (!surrounding.Any(run => IsNativeRenderableTextRun(run, paragraph))) return false;
        PdfCore.PdfParagraphStyle paragraphStyle = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
        PdfCore.PageMargins margins = GetNativeMargins(defaults.RunningContentSection, options, (0D, 0D));
        double textLeft = margins.Left + paragraphStyle.LeftIndent + paragraphStyle.FirstLineIndent;
        double textWidth = 0D;
        foreach (WordParagraph run in surrounding) {
            double? measured = fontMap.MeasureText(CreateNativeCellTextRun(run.Text, run,
                NativeTableStyleDefaults.Empty, defaults, fontMap));
            if (!measured.HasValue) return false;
            textWidth += measured.Value;
        }
        // Square wrapping is qualified here only when the entire ordinary line
        // clears the anchor horizontally. No wrapping heuristic is introduced.
        double exclusionLeft = x - (box.Anchor?.DistanceFromLeft?.Value ?? 0U) / 12700D;
        double exclusionRight = x + width + (box.Anchor?.DistanceFromRight?.Value ?? 0U) / 12700D;
        double contentWidth = defaults.RunningContentContext?.ContentWidth ?? pageSize.Width - margins.Left - margins.Right;
        if (textLeft + textWidth > margins.Left + contentWidth - paragraphStyle.RightIndent ||
            textLeft + textWidth > exclusionLeft && textLeft < exclusionRight) return false;

        if (!TryCreateNativeRunningTextBoxStyle(box, inner, defaults, fontMap, out PdfCore.PdfCanvasTextBoxStyle style)) return false;
        if (style.PaddingLeft + style.PaddingRight >= width || style.PaddingTop + style.PaddingBottom >= height) return false;
        List<PdfCore.PdfTextRun> boxRuns = CreateNativeCellParagraphRuns(inner, notes,
            NativeTableStyleDefaults.Empty, defaults, fontMap);
        if (boxRuns.Count == 0) return false;
        double boxTextWidth = 0D;
        foreach (PdfCore.PdfTextRun run in boxRuns) {
            double? measured = fontMap.MeasureText(run);
            if (!measured.HasValue) return false;
            boxTextWidth += measured.Value;
        }
        if (boxTextWidth > width - style.PaddingLeft - style.PaddingRight ||
            (style.LineHeight ?? 1.2D) * (style.FontSize ?? defaults.FontSize) *
                (innerChildren.LastOrDefault() is W.Break ? 2D : 1D) > height - style.PaddingTop - style.PaddingBottom) return false;
        var canvas = new PdfCore.PdfPageCanvas();
        canvas.TextBox(boxRuns, x, y, width, height, style);
        paragraphStyle.AnchoredCanvas = new PdfCore.PdfCanvasBlock(canvas.Items);
        prepared = new NativeAnchoredRunningTextBoxParagraph(surrounding.ToArray(), paragraphStyle);
        return true;
    }

    private static bool TryCreateNativeRunningTextBoxStyle(WordTextBox box, WordParagraph inner,
        NativeDocumentDefaults defaults, NativeFontMap fontMap, out PdfCore.PdfCanvasTextBoxStyle style) {
        var body = box.DrawingTextBodyProperties!;
        PdfCore.PdfParagraphStyle paragraph = CreateNativeParagraphStyle(inner, defaults, fontMap);
        style = new PdfCore.PdfCanvasTextBoxStyle { Background = null, BorderColor = null, BorderWidth = 0D };
        if (paragraph.SpacingBefore is > 0D || paragraph.LeftIndent != 0D || paragraph.RightIndent != 0D ||
            paragraph.FirstLineIndent != 0D) return false;
        double left = (body.LeftInset?.Value ?? 91440) / 12700D;
        double right = (body.RightInset?.Value ?? 91440) / 12700D;
        double top = (body.TopInset?.Value ?? 45720) / 12700D;
        double bottom = (body.BottomInset?.Value ?? 45720) / 12700D;
        style.Align = ResolveNativeParagraphAlign(inner, allowJustify: false);
        style.FontSize = paragraph.FontSize;
        style.LineHeight = paragraph.LineHeight;
        style.LineSpacing = paragraph.LineSpacing;
        style.VerticalAlign = PdfCore.PdfVerticalAlign.Top;
        if (left < 0D || right < 0D || top < 0D || bottom < 0D) return false;
        style.PaddingLeft = left; style.PaddingRight = right;
        style.PaddingTop = top; style.PaddingBottom = bottom;
        var properties = box.DrawingShapeProperties;
        if (properties?.GetFirstChild<A.GradientFill>() != null || properties?.GetFirstChild<A.BlipFill>() != null ||
            properties?.GetFirstChild<A.PatternFill>() != null || properties?.GetFirstChild<A.GroupFill>() != null) return false;
        W.Run? sourceRun = box.Drawing?.Ancestors<W.Run>().FirstOrDefault();
        if (sourceRun == null || box.Drawing == null) return false;
        var wordShape = new WordShape(box.Document, sourceRun.Ancestors<W.Paragraph>().First(), sourceRun, box.Drawing);
        var shape = OfficeShape.Rectangle(1D, 1D);
        shape.FillColor = null; shape.StrokeColor = null; shape.StrokeWidth = 0D;
        ApplyNativeShapeStyle(shape, wordShape, GetNativeDrawingThemeColors(GetNativeShapeSourcePart(wordShape)));
        style.Background = shape.FillColor is { } fill ? PdfCore.PdfColor.FromOfficeColor(fill) : null;
        style.BackgroundOpacity = shape.FillOpacity;
        style.BorderColor = shape.StrokeColor is { } stroke ? PdfCore.PdfColor.FromOfficeColor(stroke) : null;
        style.BorderWidth = shape.StrokeWidth;
        style.BorderDashStyle = shape.StrokeDashStyle;
        style.BorderLineCap = shape.StrokeLineCap;
        style.BorderLineJoin = shape.StrokeLineJoin;
        return !shape.StrokeOpacity.HasValue || shape.StrokeOpacity.Value == 1D;
    }
}
