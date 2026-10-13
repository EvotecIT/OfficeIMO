using DocumentFormat.OpenXml;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Validates one simple floating box and appends its paint in the resolved drawing order.</summary>
    private static bool TryAddNativeRunningTextBox(PdfCore.PdfPageCanvas canvas, WordTextBox box,
        PdfCore.PageSize pageSize, PdfCore.PageMargins margins, double textLeft, double textWidth, NativeNoteNumbering? notes,
        WordToPdfOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap) {
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
        if (!TryGetNativeRunningTextBoxFrame(box, defaults.RunningContentSection!, options, pageSize,
                out double x, out double y, out double width, out double height, out bool behindDocument)) return false;
        Wps.TextBodyProperties? body = box.DrawingTextBodyProperties;
        if (body == null || body.Rotation?.Value is > 0 or < 0 ||
            body.Vertical?.Value is { } vertical && vertical != A.TextVerticalValues.Horizontal ||
            body.Anchor?.Value is { } anchor && anchor != A.TextAnchoringTypeValues.Top ||
            body.GetFirstChild<A.NormalAutoFit>() != null) return false;

        // Each square-wrap exclusion must clear the entire ordinary line.
        // Peers may overlap each other without becoming paragraph flow blocks.
        double exclusionLeft = x - (box.Anchor?.DistanceFromLeft?.Value ?? 0U) / 12700D;
        double exclusionRight = x + width + (box.Anchor?.DistanceFromRight?.Value ?? 0U) / 12700D;
        if (textLeft + textWidth > exclusionLeft && textLeft < exclusionRight) return false;

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
        if (boxTextWidth > width - style.PaddingLeft - style.PaddingRight) return false;
        PdfCore.PdfWriter.CanvasTextBoxMeasurement? measuredBox = fontMap.MeasureCanvasTextBox(boxRuns, width, height, style);
        if (measuredBox == null || !measuredBox.TextFits) return false;
        // Running canvas paints before body content. A foreground box is safe
        // only when its visible paint clears the canonical body frame. Running
        // reservations can shrink that frame further, never enlarge it. Extend
        // its bottom to the physical page: imported exact lines may paint into
        // the footer margin and this route has no actual body-ink inventory.
        if (!behindDocument && measuredBox.PaintWidth > 0D && measuredBox.PaintHeight > 0D &&
            x + measuredBox.PaintLeft + measuredBox.PaintWidth > margins.Left &&
            x + measuredBox.PaintLeft < pageSize.Width - margins.Right &&
            y + measuredBox.PaintTop + measuredBox.PaintHeight > margins.Top &&
            y + measuredBox.PaintTop < pageSize.Height) return false;
        canvas.TextBox(boxRuns, x, y, width, height, style);
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
