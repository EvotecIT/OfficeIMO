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
        List<WordParagraph> innerRuns = GetNativeRuns(inner);
        if (HasNativeUnsupportedRunningTextBoxRunPaint(inner, innerRuns)) return false;
        OpenXmlElement[] innerChildren = innerRuns.SelectMany(run => run.EnumerateEffectiveRunContent())
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
            body.ColumnCount?.Value is > 1 ||
            body.GetFirstChild<A.PresetTextWarp>() is { } warp && warp.Preset?.Value != A.TextShapeValues.TextNoShape ||
            body.GetFirstChild<A.Scene3DType>() != null || body.GetFirstChild<A.Shape3DType>() != null ||
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
        if (!behindDocument && !defaults.BodyTextInkTopOffset.HasValue) return false;
        // Running canvas paints before body content. A foreground box is safe
        // only when its visible paint clears a conservative body envelope. Use
        // the once-measured top ink offset, the physical width for negative
        // indents, and the page bottom for imported exact-line footer overflow.
        if (!behindDocument && measuredBox.PaintWidth > 0D && measuredBox.PaintHeight > 0D &&
            x + measuredBox.PaintLeft + measuredBox.PaintWidth > 0D &&
            x + measuredBox.PaintLeft < pageSize.Width &&
            y + measuredBox.PaintTop + measuredBox.PaintHeight > margins.Top + defaults.BodyTextInkTopOffset.GetValueOrDefault() &&
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
            paragraph.FirstLineIndent != 0D || HasNativeUnsupportedRunningTextBoxParagraphPaint(inner)) return false;
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
        if (HasNativeUnsupportedRunningTextBoxPaint(properties)) return false;
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

    /// <summary>Run-only box content cannot reproduce the paragraph's panel or patterned shading.</summary>
    private static bool HasNativeUnsupportedRunningTextBoxParagraphPaint(WordParagraph paragraph) {
        W.ParagraphPropertiesBaseStyle? documentProperties = paragraph._document?._wordprocessingDocument?
            .MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?.GetFirstChild<W.ParagraphPropertiesDefault>()?
            .GetFirstChild<W.ParagraphPropertiesBaseStyle>();
        // A fill-only resolver discards auto/missing fills, which can still have a visible shading pattern.
        return HasNativeParagraphBorder(GetNativeEffectiveParagraphBorders(paragraph)) ||
            paragraph._paragraph.ParagraphProperties?.GetFirstChild<W.Shading>() != null ||
            GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId).Any(source =>
                source.GetFirstChild<W.StyleParagraphProperties>()?.GetFirstChild<W.Shading>() != null) ||
            documentProperties?.GetFirstChild<W.Shading>() != null || documentProperties?.GetFirstChild<W.ParagraphBorders>() != null;
    }

    /// <summary>Leaves paint that the fixed canvas cannot reproduce on the explicitly approximated route.</summary>
    private static bool HasNativeUnsupportedRunningTextBoxPaint(Wps.ShapeProperties? properties) {
        if (properties == null) return false;
        if (properties.GetFirstChild<A.EffectList>()?.HasChildren == true ||
            properties.GetFirstChild<A.EffectDag>()?.HasChildren == true ||
            properties.GetFirstChild<A.Scene3DType>() != null || properties.GetFirstChild<A.Shape3DType>() != null) return true;
        Wps.ShapeStyle? reference = properties.Parent?.GetFirstChild<Wps.ShapeStyle>();
        if (reference?.GetFirstChild<A.EffectReference>()?.Index?.Value is > 0U) return true;
        // Direct fills override the theme fill. The style mapper resolves direct scheme colors,
        // but does not resolve the theme's fill/line style matrix.
        if (properties.GetFirstChild<A.NoFill>() == null && properties.GetFirstChild<A.SolidFill>() == null &&
            reference?.GetFirstChild<A.FillReference>()?.Index?.Value is > 0U) return true;
        A.Outline? outline = properties.GetFirstChild<A.Outline>();
        if (outline?.GetFirstChild<A.NoFill>() != null) return false;
        if (reference?.GetFirstChild<A.LineReference>()?.Index?.Value is > 0U) return true;
        if (outline == null) return false;
        return outline.GetFirstChild<A.GradientFill>() != null || outline.GetFirstChild<A.PatternFill>() != null ||
            outline.GetFirstChild<A.CustomDash>() != null ||
            outline.GetFirstChild<A.PresetDash>()?.Val?.Value is { } dash && dash != A.PresetLineDashValues.Solid ||
            outline.CompoundLineType?.Value is { } compound && compound != A.CompoundLineValues.Single ||
            outline.Alignment?.Value is { } alignment && alignment != A.PenAlignmentValues.Center ||
            outline.CapType?.Value is { } cap && cap != A.LineCapValues.Flat ||
            outline.GetFirstChild<A.Round>() != null || outline.GetFirstChild<A.LineJoinBevel>() != null ||
            outline.GetFirstChild<A.Miter>()?.Limit != null ||
            outline.GetFirstChild<A.HeadEnd>()?.Type?.Value is { } head && head != A.LineEndValues.None ||
            outline.GetFirstChild<A.TailEnd>()?.Type?.Value is { } tail && tail != A.LineEndValues.None;
    }
}
