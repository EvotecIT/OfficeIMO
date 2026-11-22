using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using WP14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>The page rectangle, insets and stacking of an anchored header or footer text box.</summary>
    private sealed record NativePlacedTextBoxFrame(double X, double Y, double Width, double Height,
        double LeftInset, double TopInset, double RightInset, double BottomInset,
        bool ParagraphRelative, double? MarginOrigin, bool BehindDocument, long ZOrder,
        PdfCore.PdfImageTextExclusion? Exclusion);

    /// <summary>
    /// Lays out anchored header and footer text boxes as Word does: each box keeps its own position, size
    /// and insets, and its paragraphs use the ordinary paragraph route inside that rectangle. The anchor
    /// paragraph keeps its own text and line.
    /// </summary>
    private static bool TryRenderNativePlacedRunningTextBoxParagraph(INativePdfFlow pdf, WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, (int Level, string Marker)? marker, Func<WordParagraph, (int Level, string Marker)?> getMarker,
        NativeNoteNumbering notes, WordPdfExportOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap,
        WordParagraph? nextParagraph) {
        if (defaults.RunningContentContext == null || defaults.RunningContentSection is not WordSection section ||
            !TryCollectNativePlacedRunningTextBoxes(pdf.PageSize, paragraph, runs, section, options,
                out List<(WordTextBox Box, NativePlacedTextBoxFrame Frame)> boxes, out List<WordParagraph> own, defaults)) return false;
        var anchored = new PdfCore.PdfPageCanvas();
        foreach (var (box, frame) in boxes.OrderBy(entry => entry.Frame.ZOrder)) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            // The box paints in its own coordinates, from its top-left corner.
            var content = new PdfCore.PdfPageCanvas(allowOutOfPageCoordinates: true);
            OfficeShape? decoration = CreateNativePlacedTextBoxDecoration(box, frame);
            if (decoration != null)
                content.Shape(decoration, global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(0D));
            double outline = GetNativePlacedTextBoxOutlineInset(decoration);
            double contentWidth = frame.Width - frame.LeftInset - frame.RightInset - 2D * outline;
            double contentHeight = frame.Height - frame.TopInset - frame.BottomInset - 2D * outline;
            IReadOnlyList<WordParagraph> inner = GetNativeTextBoxParagraphs(box);
            if (inner.Count > 0 && contentWidth > 0D && contentHeight > 0D) {
                IReadOnlyList<PdfCore.IPdfBlock> blocks = pdf.CaptureBlocks(builder => RenderNativeTextBoxFrameContent(builder,
                    pdf.PageSize, section, inner, getMarker, notes, options, defaults, fontMap, contentWidth));
                if (blocks.Count > 0)
                    content.PlacedFlow(blocks, frame.LeftInset + outline, frame.TopInset + outline, contentWidth, contentHeight,
                        0D, 0D, frame.Width, frame.Height);
            }
            IReadOnlyList<PdfCore.PdfCanvasItem> items = content.Items;
            // An unpainted box still keeps wrapped text out of its rectangle.
            if (items.Count == 0 && frame.Exclusion == null) continue;
            if (frame.BehindDocument && frame.Exclusion != null)
                items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasBehindTextItem(items, frame.ZOrder) };
            if (frame.Exclusion != null) {
                // A wrapping box shares the picture route, which places it once and keeps text out of it.
                anchored.AddItems(new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasAnchoredImageItem(box.Drawing!,
                    new PdfCore.PdfCanvasBlock(items), frame.Width, frame.Height,
                    new PdfCore.PdfAnchoredImagePosition(frame.X, frame.Y, frame.MarginOrigin.HasValue, frame.MarginOrigin ?? 0D,
                        frame.ParagraphRelative), frame.Exclusion) });
                continue;
            }
            var placed = new PdfCore.PdfPageCanvas();
            placed.Effect(OfficeTransform.Translate(frame.X, frame.Y), 1D, nested => nested.AddItems(items));
            items = placed.Items;
            if (frame.MarginOrigin is { } marginOrigin)
                items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasMarginAnchorItem(items, marginOrigin) };
            if (frame.ParagraphRelative)
                items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasParagraphAnchorItem(items) };
            // Word stacks the box against all of the story's text: behind it, or in front of it.
            items = frame.BehindDocument
                ? new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasBehindTextItem(items, frame.ZOrder) }
                : new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasInFrontOfTextItem(items, frame.ZOrder) };
            anchored.AddItems(items);
        }

        PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph, defaults, fontMap);
        bool hasMarker = marker is { Marker.Length: > 0 };
        if (hasMarker) {
            ApplyNativeInlineListIndent(paragraph, style);
            ApplyNativeInlineListMarkerAlignment(paragraph, marker!.Value.Marker, style, defaults, fontMap);
        }
        if (ShouldSuppressNativeContextualSpacingAfter(paragraph, nextParagraph)) style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
        if (ShouldSuppressNativeContextualSpacingBefore(paragraph)) style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);
        if (anchored.Items.Count > 0) style.AnchoredCanvas = new PdfCore.PdfCanvasBlock(anchored.Items);
        WordParagraph[] selected = own.ToArray();
        bool visible = selected.Any(run => IsNativeRenderableTextRun(run, paragraph));
        List<int> references = GetNativeMixedTextBoxNoteNumbers(selected, notes, new HashSet<NativeNoteKey>());
        global::OfficeIMO.OfficeAlignment align = ResolveNativeParagraphAlign(paragraph, allowJustify: false);
        Action<PdfCore.PdfParagraph> build;
        if (visible || references.Count > 0 || hasMarker) {
            string text = string.Concat(selected.Select(run => run.Text));
            PdfCore.PdfParagraphStyle markerStyle = style;
            build = builder => AddNativeParagraphContent(builder, paragraph, marker, selected, visible, text,
                references, notes, options, defaults, fontMap,
                inlineMarkerColumnWidth: hasMarker ? ResolveNativeInlineListMarkerColumnWidth(paragraph, marker!.Value.Marker, markerStyle, defaults, fontMap) : null,
                projectedContentOnly: true);
        } else {
            // The anchor paragraph still owns an empty line in its own font.
            double fontSize = ResolveNativeParagraphFontSize(paragraph, defaults, GetNativeParagraphStyleDefaults(paragraph));
            build = builder => builder.AddLineBreak().SetFontSize(global::OfficeIMO.OfficeLength.Points(fontSize));
        }
        // The anchor paragraph keeps its own shading, borders and bottom rule, as on the ordinary route.
        if (CreateNativeParagraphPanelStyle(paragraph, style) is { } panelStyle) {
            style = JoinNativeAdjacentParagraphShading(nextParagraph, style, panelStyle, defaults).Clone();
            style.LeftIndent -= panelStyle.LeftInset;
            style.RightIndent -= panelStyle.RightInset;
            pdf.PanelParagraph(builder => {
                using NativePaintedPanelScope paintedPanel = EnterNativePaintedPanel(paragraph, panelStyle.Background);
                build(builder);
            }, panelStyle, align, null, style);
        } else if (!hasMarker && CreateNativeBottomBorderRuleStyle(paragraph, style) is { } ruleStyle) {
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
            pdf.Paragraph(build, align, style: style);
            pdf.HR(style: ruleStyle);
        } else {
            pdf.Paragraph(build, align, style: style);
        }
        RecordNativeParagraphBarTabDiagnostics(paragraph, options, "anchored text-box paragraph");
        return true;
    }

    /// <summary>
    /// Lays out the box's paragraphs exactly as the story lays out its own: the same paragraph route,
    /// spacing and contextual-spacing rules, markers, shading and borders, at the box's content width.
    /// </summary>
    private static void RenderNativeTextBoxFrameContent(PdfCore.PdfContentBuilder builder, PdfCore.PdfPageSize pageSize,
        WordSection section, IReadOnlyList<WordParagraph> paragraphs, Func<WordParagraph, (int Level, string Marker)?> getMarker,
        NativeNoteNumbering notes, WordPdfExportOptions? options, NativeDocumentDefaults defaults, NativeFontMap fontMap,
        double contentWidth) {
        var flow = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(builder, pageSize), preserveInitialSpacingBefore: true);
        var destinations = new Dictionary<W.Paragraph, string>();
        for (int index = 0; index < paragraphs.Count; index++) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            RenderNativeElement(flow, paragraphs[index], section, getMarker, Array.Empty<int>(), notes, options,
                Array.Empty<NativeTableOfContentsEntry>(), destinations, contentWidth, defaults, fontMap,
                index + 1 < paragraphs.Count ? paragraphs[index + 1] : null);
        }
    }

    /// <summary>
    /// Qualifies every visible box of the paragraph and selects the paragraph's own content. With the
    /// document defaults it also checks the paint the frame reproduces and the body text it must clear.
    /// </summary>
    private static bool TryCollectNativePlacedRunningTextBoxes(PdfCore.PdfPageSize pageSize, WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs, WordSection section, WordPdfExportOptions? options,
        out List<(WordTextBox Box, NativePlacedTextBoxFrame Frame)> boxes, out List<WordParagraph> own,
        NativeDocumentDefaults? defaults = null) {
        boxes = new List<(WordTextBox, NativePlacedTextBoxFrame)>();
        own = new List<WordParagraph>();
        // A framed anchor paragraph is positioned by its own frame; keep the bounded flow for it.
        if (WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph).Count > 0 ||
            paragraph._paragraph.ParagraphProperties?.FrameProperties != null) return false;
        // Boxes that must not overlap would need Word's repositioning; keep the bounded flow.
        if (defaults?.RunningStoryHasDisallowedTextBoxPeers == true) return false;
        foreach (WordParagraph run in runs) {
            if (IsNativeHiddenTextRun(run, paragraph)) continue;
            foreach (OpenXmlElement child in run.EnumerateEffectiveRunContent()) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (child is W.RunProperties) continue;
                WordParagraph view = CreateNativeRunContentView(run, new[] { child });
                List<WordTextBox> childBoxes = view.GetTextBoxes().ToList();
                if (childBoxes.Count > 0) {
                    if (view.GetGroupTextBoxes().Any()) return false;
                    foreach (WordTextBox box in childBoxes) {
                        if (!TryGetNativePlacedTextBoxFrame(box, section, options, pageSize, out NativePlacedTextBoxFrame? frame) ||
                            !CanPaintNativePlacedTextBox(box, frame!))
                            return false;
                        boxes.Add((box, frame!));
                    }
                    continue;
                }
                // Other drawables keep their established routes.
                if (view.GetGroupTextBoxes().Any() || view.IsImage || view.Shape != null || view.Chart != null ||
                    HasNativeParagraphShapeGroups(new[] { view })) return false;
                own.Add(view);
            }
        }
        return boxes.Count > 0;
    }

    /// <summary>Resolves an anchored rectangular text box to its page rectangle without modifying source XML.</summary>
    private static bool TryGetNativePlacedTextBoxFrame(WordTextBox box, WordSection section, WordPdfExportOptions? options,
        PdfCore.PdfPageSize page, out NativePlacedTextBoxFrame? frame) {
        frame = null;
        var properties = box.DrawingShapeProperties;
        // Nested boxes keep the bounded flow; a frame lays out ordinary paragraphs only.
        if (box.Drawing == null || box.Content == null || box.Content.Descendants<W.TextBoxContent>().Any() ||
            box.Content.Descendants<W.Table>().Any() ||
            properties?.GetFirstChild<A.CustomGeometry>() != null ||
            properties?.GetFirstChild<A.PresetGeometry>() is { } geometry &&
                !string.Equals(geometry.Preset?.InnerText, "rect", StringComparison.Ordinal) ||
            properties?.GetFirstChild<A.Transform2D>() is { } transform &&
                (transform.Rotation?.Value is > 0 or < 0 || transform.HorizontalFlip?.Value == true || transform.VerticalFlip?.Value == true))
            return false;
        if (!WordDrawingLayoutReader.TryRead(box.Drawing, out WordDrawingLayoutInfo layout) ||
            layout.Placement != WordDrawingPlacementKind.Anchored || layout.UsesSimplePosition) return false;
        var body = box.DrawingTextBodyProperties;
        if (body != null && (body.Rotation?.Value is > 0 or < 0 ||
            body.Vertical?.Value is { } vertical && vertical != A.TextVerticalValues.Horizontal ||
            body.ColumnCount?.Value is > 1)) return false;

        string? horizontalReference = layout.HorizontalRelativeFrom == "column" ? "margin" : layout.HorizontalRelativeFrom;
        string? verticalReference = layout.VerticalRelativeFrom == "line" ? "paragraph" : layout.VerticalRelativeFrom;
        PdfCore.PdfPageMargins margins = GetNativeMargins(section, options, (0D, 0D));
        double width = layout.WidthPoints, height = layout.HeightPoints;
        WP14.RelativeWidth? relativeWidth = box.Anchor?.GetFirstChild<WP14.RelativeWidth>();
        WP14.RelativeHeight? relativeHeight = box.Anchor?.GetFirstChild<WP14.RelativeHeight>();
        if (!TryGetNativeRunningRelativeExtent(relativeWidth?.PercentageWidth?.Text, relativeWidth?.ObjectId?.InnerText,
                page.Width.ToPoints(), margins.Left.ToPoints(), margins.Right.ToPoints(), ref width) ||
            !TryGetNativeRunningRelativeExtent(relativeHeight?.PercentageHeight?.Text, relativeHeight?.RelativeFrom?.InnerText,
                page.Height.ToPoints(), margins.Top.ToPoints(), margins.Bottom.ToPoints(), ref height) ||
            !(width > 0D) || !(height > 0D)) return false;
        if (!TryGetNativeGroupAxis(horizontalReference, layout.HorizontalOffsetPoints, layout.HorizontalAlignment,
                page.Width.ToPoints(), margins.Left.ToPoints(), margins.Right.ToPoints(), width, false, out double x) ||
            !TryGetNativeGroupAxis(verticalReference, layout.VerticalOffsetPoints, layout.VerticalAlignment,
                page.Height.ToPoints(), margins.Top.ToPoints(), margins.Bottom.ToPoints(), height, true, out double y)) return false;

        if (box.Anchor is not DW.Anchor anchor || !TryReadNativePlacedTextBoxExclusion(anchor, out PdfCore.PdfImageTextExclusion? exclusion))
            return false;
        // Word keeps text clear of the box's effect extent as well as its wrap distances.
        if (exclusion != null && anchor.GetFirstChild<DW.EffectExtent>() is { } extent) {
            static double Points(Int64Value? value) => Math.Max(0L, value?.Value ?? 0L) / 12700D;
            exclusion = new PdfCore.PdfImageTextExclusion(exclusion.Kind, exclusion.Side,
                exclusion.DistanceLeft + Points(extent.LeftEdge), exclusion.DistanceTop + Points(extent.TopEdge),
                exclusion.DistanceRight + Points(extent.RightEdge), exclusion.DistanceBottom + Points(extent.BottomEdge));
        }
        double Inset(Int32Value? value, int fallback) => (value?.Value ?? fallback) / 12700D;
        double left = Inset(body?.LeftInset, 91440), right = Inset(body?.RightInset, 91440);
        double top = Inset(body?.TopInset, 45720), bottom = Inset(body?.BottomInset, 45720);
        if (left < 0D || right < 0D || top < 0D || bottom < 0D) return false;
        var placed = new NativePlacedTextBoxFrame(x, y, width, height, left, top, right, bottom,
            ParagraphRelative: verticalReference == "paragraph",
            MarginOrigin: horizontalReference == "margin" ? margins.Left.ToPoints() : null,
            BehindDocument: layout.BehindDocument, ZOrder: layout.RelativeHeight, Exclusion: exclusion);
        // Qualify the paragraphs at the width rendering gives them, inside the insets and the outline.
        double outline = GetNativePlacedTextBoxOutlineInset(CreateNativePlacedTextBoxDecoration(box, placed));
        if (!CanLayOutNativePlacedTextBoxContent(box, width - left - right - 2D * outline)) return false;
        frame = placed;
        return true;
    }

    /// <summary>
    /// A frame lays out ordinary paragraphs on one page at a fixed width. Page and column breaks, floating
    /// frames, floating drawings and indents that leave no text width keep the bounded flow instead.
    /// </summary>
    private static bool CanLayOutNativePlacedTextBoxContent(WordTextBox box, double contentWidth) {
        W.TextBoxContent content = box.Content!;
        if (content.Descendants<W.Break>().Any(item => item.Type?.Value == W.BreakValues.Page || item.Type?.Value == W.BreakValues.Column) ||
            content.Descendants<W.FrameProperties>().Any() || content.Descendants<W.SectionProperties>().Any() ||
            content.Descendants<DW.Anchor>().Any() || content.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Any() ||
            content.Descendants<DocumentFormat.OpenXml.Vml.Group>().Any()) return false;
        foreach (WordParagraph inner in GetNativeTextBoxParagraphs(box)) {
            if (HasNativePageBreakBefore(inner)) return false;
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(inner);
            if (WordDocumentTraversal.GetListInfo(inner).HasValue) ApplyNativeInlineListIndent(inner, style);
            double textWidth = contentWidth - style.LeftIndent.ToPoints() - style.RightIndent.ToPoints();
            double firstLineWidth = textWidth - style.FirstLineIndent.ToPoints();
            if (!(textWidth > 0D) || !(firstLineWidth > 0D)) return false;
        }
        return true;
    }

    /// <summary>A text box is rectangular, so tight and through wrapping clear the same rectangle as square wrapping.</summary>
    private static bool TryReadNativePlacedTextBoxExclusion(DW.Anchor anchor, out PdfCore.PdfImageTextExclusion? exclusion) {
        exclusion = null;
        if (anchor.GetFirstChild<DW.WrapTight>() == null && anchor.GetFirstChild<DW.WrapThrough>() == null)
            return TryReadNativeAnchoredImageExclusion(anchor, out exclusion, out _);
        OpenXmlElement wrapping = (OpenXmlElement?)anchor.GetFirstChild<DW.WrapTight>() ?? anchor.GetFirstChild<DW.WrapThrough>()!;
        PdfCore.PdfImageTextWrapSide side = ReadNativeImageWrapAttribute(wrapping, "wrapText") switch {
            "left" => PdfCore.PdfImageTextWrapSide.Left,
            "right" => PdfCore.PdfImageTextWrapSide.Right,
            "largest" => PdfCore.PdfImageTextWrapSide.Largest,
            _ => PdfCore.PdfImageTextWrapSide.BothSides
        };
        exclusion = new PdfCore.PdfImageTextExclusion(PdfCore.PdfImageTextWrapKind.Square, side,
            (anchor.DistanceFromLeft?.Value ?? 0U) / 12700D, (anchor.DistanceFromTop?.Value ?? 0U) / 12700D,
            (anchor.DistanceFromRight?.Value ?? 0U) / 12700D, (anchor.DistanceFromBottom?.Value ?? 0U) / 12700D);
        return true;
    }

    /// <summary>
    /// The frame paints solid fills and plain outlines. Like Word, the header and footer layer paints
    /// before the body, so a box that reaches into the body area stays under the body text even when
    /// it is in front of the header text.
    /// </summary>
    private static bool CanPaintNativePlacedTextBox(WordTextBox box, NativePlacedTextBoxFrame frame) {
        var properties = box.DrawingShapeProperties;
        if (properties?.GetFirstChild<A.GradientFill>() != null || properties?.GetFirstChild<A.BlipFill>() != null ||
            properties?.GetFirstChild<A.PatternFill>() != null || properties?.GetFirstChild<A.GroupFill>() != null ||
            HasNativeUnsupportedRunningTextBoxPaint(properties)) return false;
        if (CreateNativePlacedTextBoxDecoration(box, frame) is { StrokeOpacity: { } opacity } && opacity != 1D) return false;
        if (box.DrawingTextBodyProperties is { } body &&
            (body.Anchor?.Value is { } anchor && anchor != A.TextAnchoringTypeValues.Top ||
             body.GetFirstChild<A.PresetTextWarp>() is { } warp && warp.Preset?.Value != A.TextShapeValues.TextNoShape ||
             body.GetFirstChild<A.Scene3DType>() != null || body.GetFirstChild<A.Shape3DType>() != null ||
             body.GetFirstChild<A.NormalAutoFit>() != null)) return false;
        return true;
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

    private static bool TryGetNativeRunningRelativeExtent(string? percentage, string? reference,
        double pageLength, double leadingMargin, double trailingMargin, ref double extent) {
        if (percentage == null) return true;
        if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(percentage, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
            double.IsNaN(value) || double.IsInfinity(value) || value < 0D) return false;
        if (value == 0D) return true;
        if (reference != "page" && reference != "margin") return false;
        extent = (reference == "page" ? pageLength : pageLength - leadingMargin - trailingMargin) * value / 100000D;
        return !double.IsNaN(extent) && !double.IsInfinity(extent);
    }

    /// <summary>Word sets the text inside the outline: half the line width beyond each inset.</summary>
    private static double GetNativePlacedTextBoxOutlineInset(OfficeShape? decoration) =>
        decoration?.StrokeColor != null ? decoration.StrokeWidth / 2D : 0D;

    /// <summary>The box's own fill and outline, painted under its text.</summary>
    private static OfficeShape? CreateNativePlacedTextBoxDecoration(WordTextBox box, NativePlacedTextBoxFrame frame) {
        W.Run? sourceRun = box.Drawing?.Ancestors<W.Run>().FirstOrDefault();
        W.Paragraph? sourceParagraph = sourceRun?.Ancestors<W.Paragraph>().FirstOrDefault();
        if (sourceRun == null || sourceParagraph == null || box.Drawing == null) return null;
        var wordShape = new WordShape(box.Document, sourceParagraph, sourceRun, box.Drawing);
        var shape = OfficeShape.Rectangle(frame.Width, frame.Height);
        shape.FillColor = null; shape.StrokeColor = null; shape.StrokeWidth = 0D;
        ApplyNativeShapeStyle(shape, wordShape, GetNativeDrawingThemeColors(GetNativeShapeSourcePart(wordShape)));
        return shape.FillColor != null || shape.StrokeColor != null && shape.StrokeWidth > 0D ? shape : null;
    }
}
