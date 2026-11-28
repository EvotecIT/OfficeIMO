using System.Globalization;
using OfficeIMO.Drawing;
using static OfficeIMO.Word.OpenXml.WordXml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const string NativeWordGroupNamespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

        private static bool HasNativeParagraphShapeGroups(IReadOnlyList<WordParagraph> runs) =>
            runs.Any(run => run.EffectiveRunContentNodes().SelectMany(child => child.DescendantElements().Prepend(child))
                .Any(child => child.Is(V, "group") || child.NamespaceUri == NativeWordGroupNamespace && child.LocalName == "wgp"));

        /// <summary>Projects a group as one object, retaining its child coordinate system and paragraph anchor.</summary>
        private static bool RenderNativeParagraphShapeGroups(INativePdfFlow pdf, WordParagraph paragraph,
            IReadOnlyList<WordParagraph> runs, global::OfficeIMO.OfficeAlignment align, WordPdfExportOptions? options, PdfCore.PdfParagraphStyle style, ref int imageCount, NativeObjectParagraphSpacing? paragraphSpacing,
            bool runningStory = false) =>
            ProjectNativeParagraphShapeGroups(pdf.PageSize, paragraph, runs, options, style,
                scene => RenderNativeFlowObject(pdf, paragraphSpacing,
                    flow => flow.Drawing(scene, align, spacingAfter: 0D)), ref imageCount, runningStory);

        private static bool HasNativeParagraphShapeGroupAnchor(PdfCore.PdfPageSize pageSize, WordParagraph paragraph,
            IReadOnlyList<WordParagraph> runs, WordPdfExportOptions? options) {
            var style = new PdfCore.PdfParagraphStyle();
            int imageCount = 0;
            // Keep cancellation, image limits and resource policy, while keeping
            // this read-only preparation's warnings out of the real operation.
            ProjectNativeParagraphShapeGroups(pageSize, paragraph, runs, options?.CloneForConversion(), style,
                null, ref imageCount);
            return style.AnchoredCanvas != null;
        }

        // Both the renderer and the pre-report classification prepare the same
        // supported canvas. Only the renderer supplies a flow-object callback.
        private static bool ProjectNativeParagraphShapeGroups(PdfCore.PdfPageSize pageSize, WordParagraph paragraph,
            IReadOnlyList<WordParagraph> runs, WordPdfExportOptions? options, PdfCore.PdfParagraphStyle style,
            Action<OfficeDrawing>? renderFlowObject, ref int imageCount, bool runningStory = false) {
            bool renderedFlowObject = false;
            foreach (NativeShapeGroupSource source in EnumerateNativeVisibleShapeGroups(paragraph, runs)) {
                if (source.Legacy.Exists) {
                    RenderNativeLegacyBodyGroup(pageSize, paragraph, source.Legacy, options, style, ref imageCount);
                    continue;
                }
                WordParagraph run = source.Run;
                OfficeXmlNode drawing = source.Drawing;
                OfficeXmlNode group = GetNativeDrawingShapeGroup(drawing);
                if (group.IsNone || !WordConverterNodes.TryReadDrawingLayout(drawing, out WordDrawingLayoutInfo layout)) continue;
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (layout.WidthPoints <= 0D || layout.HeightPoints <= 0D) {
                    WarnNativeGroup(options, "native-shape-group-unsupported", "The shape group has no renderable extent.");
                    continue;
                }

                var scene = new OfficeDrawing(layout.WidthPoints, layout.HeightPoints);
                var theme = GetNativeDrawingThemeColors(group);
                bool native = TryAddNativeGroupChildren(scene, group, 0D, 0D, scene.Width, scene.Height, run, drawing, theme, options, 0);
                if (layout.Placement == WordDrawingPlacementKind.Inline) {
                    if (native) {
                        renderFlowObject?.Invoke(scene);
                        renderedFlowObject = true;
                    }
                    else WarnNativeGroup(options, "native-shape-group-unsupported", "The inline shape group contains unsupported geometry or transforms.");
                    continue;
                }

                // A header or footer story stacks its in-front groups over its own text, as Word does.
                if (!TryGetNativeGroupPosition(pageSize, paragraph, layout, options, out double x, out double y,
                        out bool paragraphRelative, out double? horizontalMarginOrigin, allowInFrontOfText: runningStory)) {
                    if (native) {
                        renderFlowObject?.Invoke(scene);
                        renderedFlowObject = true;
                        WarnNativeGroup(options, "native-shape-group-flowed", "The shape group's anchor is outside the fixed-placement contract; it was placed in document flow.");
                    } else WarnNativeGroup(options, "native-shape-group-unsupported", "The shape group's geometry and anchor could not be mapped.");
                    continue;
                }

                var canvas = new PdfCore.PdfPageCanvas();
                if (native) canvas.Drawing(scene, global::OfficeIMO.OfficeLength.Points(x), global::OfficeIMO.OfficeLength.Points(y), global::OfficeIMO.OfficeLength.Points(scene.Width), global::OfficeIMO.OfficeLength.Points(scene.Height));
                else {
                    OfficeXmlNode choice = drawing.Ancestor(Mc, "Choice");
                    OfficeXmlNode fallbackBranch = choice.IsNone ? default : choice.Parent.Child(Mc, "Fallback");
                    OfficeXmlNode fallback = fallbackBranch.IsNone ? default : fallbackBranch.FirstDescendantElement(V, "group");
                    if (fallback.IsNone) {
                        WarnNativeGroup(options, "native-shape-group-unsupported", "The shape group contains unsupported geometry or transforms and has no VML fallback.");
                        continue;
                    }
                    CountNativeVmlGroupImages(fallback, options, ref imageCount);
                    (double cw, double ch) = GetNativeVmlCoordSize(fallback, scene.Width, scene.Height);
                    (double cx, double cy) = GetNativeVmlCoordOrigin(fallback);
                    var frame = new NativeVmlFrame(0D, 0D, scene.Width, scene.Height, cw, ch, cx, cy);
                    bool rendered = false;
                    canvas.Effect(OfficeTransform.Translate(x, y), 1D, target =>
                        rendered = RenderNativeVmlCoverChildren(target, paragraph._document, fallback.Elements(),
                            frame, pageSize.Width.ToPoints(), pageSize.Height.ToPoints(), options?.CancellationToken ?? default));
                    if (!rendered) {
                        WarnNativeGroup(options, "native-shape-group-unsupported", "The shape group's VML fallback produced no visible content.");
                        continue;
                    }
                    WarnNativeGroup(options, "native-shape-group-vml-fallback", "The shape group was rendered through its VML fallback because its DrawingML geometry or transforms are unsupported.");
                }
                AddNativeGroupCanvas(style, canvas, paragraphRelative, layout.RelativeHeight, horizontalMarginOrigin,
                    inFrontOfText: !layout.BehindDocument);
            }
            return renderedFlowObject;
        }

        private readonly record struct NativeShapeGroupSource(WordParagraph Run, OfficeXmlNode Drawing, OfficeXmlNode Legacy);

        /// <summary>The <c>wpg:wgp</c> a drawing's graphic data holds, or <see cref="OfficeXmlNode.None"/>.</summary>
        private static OfficeXmlNode GetNativeDrawingShapeGroup(OfficeXmlNode drawing) {
            OfficeXmlNode graphicData = drawing.FirstDescendantSkippingExtLst(A, "graphicData");
            return graphicData.IsNone ? default : graphicData.Elements()
                .FirstOrDefault(child => child.NamespaceUri == NativeWordGroupNamespace && child.LocalName == "wgp");
        }

        /// <summary>
        /// The shape groups Word draws for a paragraph: DrawingML groups and outermost VML groups of
        /// visible runs, from the markup-compatibility branch Word selects, that are not hidden
        /// themselves. Counting, eligibility and painting all read this one list.
        /// </summary>
        private static IEnumerable<NativeShapeGroupSource> EnumerateNativeVisibleShapeGroups(WordParagraph paragraph,
            IEnumerable<WordParagraph> runs) {
            foreach (WordParagraph run in runs) {
                if (IsNativeHiddenTextRun(run, paragraph)) continue;
                foreach (OfficeXmlNode content in run.EffectiveRunContentNodes()) {
                    if (content.IsW("drawing")) {
                        if (IsNativeShapeGroupDrawing(content) && !IsNativeHiddenShapeGroupDrawing(content))
                            yield return new NativeShapeGroupSource(run, content, default);
                        continue;
                    }
                    IEnumerable<OfficeXmlNode> legacyGroups = content.DescendantElements(V, "group");
                    if (content.Is(V, "group")) legacyGroups = legacyGroups.Prepend(content);
                    foreach (OfficeXmlNode legacy in legacyGroups) {
                        if (legacy.Ancestor(V, "group").Exists || IsNativeVmlHidden(legacy) || !IsNativeSelectedBranchContent(legacy, content) ||
                            legacy.Ancestor(W, "p") != paragraph.Node()) continue;
                        yield return new NativeShapeGroupSource(run, default, legacy);
                    }
                }
            }
        }

        /// <summary>A DrawingML group is hidden through its drawing's properties or the group's own non-visual properties.</summary>
        private static bool IsNativeHiddenShapeGroupDrawing(OfficeXmlNode drawing) {
            OfficeXmlNode anchor = drawing.Child(Wp, "anchor");
            OfficeXmlNode properties = anchor.IsNone ? default : anchor.Child(Wp, "docPr");
            if (properties.IsNone && drawing.Child(Wp, "inline") is { IsNone: false } inline) properties = inline.Child(Wp, "docPr");
            if (ParseBoolean(properties.Attr("hidden")) == true) return true;
            OfficeXmlNode group = GetNativeDrawingShapeGroup(drawing);
            return group.Exists && group.Elements().FirstOrDefault(child => child.LocalName == "cNvPr") is { IsNone: false } groupProperties &&
                ReadNativeOnOffAttribute(groupProperties, "hidden");
        }

        private static bool ReadNativeOnOffAttribute(OfficeXmlNode element, string name) =>
            element.Attr(name) is "1" or "true" or "on";

        /// <summary>Every markup-compatibility branch between the element and the run content is the one Word selects.</summary>
        private static bool IsNativeSelectedBranchContent(OfficeXmlNode element, OfficeXmlNode stop) {
            OfficeXmlNode stopParent = stop.Parent;
            for (OfficeXmlNode node = element.Parent; node.Exists && node != stopParent; node = node.Parent) {
                if ((node.Is(Mc, "Choice") || node.Is(Mc, "Fallback")) && node.Parent is { IsNone: false } alternate &&
                    alternate.Is(Mc, "AlternateContent") && WordConverterNodes.SelectAlternateContentBranch(alternate) != node) return false;
            }
            return true;
        }

        private static void AddNativeGroupCanvas(PdfCore.PdfParagraphStyle style, PdfCore.PdfPageCanvas canvas,
            bool paragraphRelative, long zOrder, double? horizontalMarginOrigin = null, bool inFrontOfText = false) {
            var combined = new PdfCore.PdfPageCanvas();
            if (style.AnchoredCanvas != null) combined.AddItems(style.AnchoredCanvas.Items);
            IReadOnlyList<PdfCore.PdfCanvasItem> items = canvas.Items;
            if (horizontalMarginOrigin.HasValue)
                items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasMarginAnchorItem(items, horizontalMarginOrigin.Value) };
            if (paragraphRelative)
                items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasParagraphAnchorItem(items) };
            combined.AddItems(new PdfCore.PdfCanvasItem[] { inFrontOfText
                ? new PdfCore.PdfCanvasInFrontOfTextItem(items, zOrder)
                : new PdfCore.PdfCanvasBehindTextItem(items, zOrder) });
            style.AnchoredCanvas = new PdfCore.PdfCanvasBlock(combined.Items);
        }

        private static void RenderNativeLegacyBodyGroup(PdfCore.PdfPageSize pageSize, WordParagraph paragraph, OfficeXmlNode group,
            WordPdfExportOptions? options, PdfCore.PdfParagraphStyle style, ref int imageCount) {
            Dictionary<string, string> vmlStyle = ParseNativeVmlStyle(group.Attr("style"));
            if (IsNativeVmlHidden(group)) return;
            long zIndex = 0;
            bool supportedLayer = vmlStyle.TryGetValue("z-index", out string? z) &&
                long.TryParse(z, NumberStyles.Integer, CultureInfo.InvariantCulture, out zIndex) && zIndex < 0;
            bool pageX = vmlStyle.TryGetValue("mso-position-horizontal-relative", out string? horizontal) && horizontal == "page";
            bool pageY = vmlStyle.TryGetValue("mso-position-vertical-relative", out string? vertical) && vertical == "page";
            if (!supportedLayer || !pageX || (!pageY && vertical != null && vertical != "text")) {
                WarnNativeGroup(options, "native-shape-group-unsupported", "The legacy shape group's anchor is outside the non-wrapping behind-text placement contract.");
                return;
            }
            CountNativeVmlGroupImages(group, options, ref imageCount);
            var frame = new NativeVmlFrame(0D, 0D, pageSize.Width.ToPoints(), pageSize.Height.ToPoints(),
                pageSize.Width.ToPoints(), pageSize.Height.ToPoints(), 0D, 0D);
            var canvas = new PdfCore.PdfPageCanvas();
            if (RenderNativeVmlGroup(canvas, paragraph._document, group, frame, pageSize.Width.ToPoints(), pageSize.Height.ToPoints(), options?.CancellationToken ?? default))
                AddNativeGroupCanvas(style, canvas, paragraphRelative: !pageY, zOrder: zIndex);
            else WarnNativeGroup(options, "native-shape-group-unsupported", "The legacy shape group produced no visible content.");
        }

        private static void CountNativeVmlGroupImages(OfficeXmlNode group, WordPdfExportOptions? options, ref int imageCount) {
            int imageLimit = options?.MaxImagesPerParagraph ?? 1_000;
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(imageLimit, nameof(WordPdfExportOptions.MaxImagesPerParagraph));
            foreach (OfficeXmlNode image in group.DescendantElements(V, "imagedata")) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (image.Ancestors().Any(ancestor => (ancestor.Is(V, "shape") || ancestor.Is(V, "group")) &&
                        IsNativeVmlHidden(ancestor))) continue;
                if (++imageCount > imageLimit)
                    throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
            }
        }

        private static void WarnNativeGroup(WordPdfExportOptions? options, string code, string message) {
            if (options != null) AddNativeExportWarning(options, code, "body paragraph shape group", message);
        }

        private static bool TryGetNativeGroupPosition(PdfCore.PdfPageSize pageSize, WordParagraph paragraph,
            WordDrawingLayoutInfo layout, WordPdfExportOptions? options, out double x, out double y,
            out bool paragraphRelative, out double? horizontalMarginOrigin, bool allowInFrontOfText = false) {
            x = y = 0D;
            horizontalMarginOrigin = null;
            paragraphRelative = layout.VerticalRelativeFrom == "paragraph";
            if (layout.Wrap != WordDrawingWrapKind.None || !layout.BehindDocument && !allowInFrontOfText) return false;
            PdfCore.PdfPageMargins margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(0D));
            if (layout.HorizontalRelativeFrom == "margin" || layout.VerticalRelativeFrom == "margin") {
                if (paragraph.Parent is not WordSection section) return false;
                margins = GetNativeMargins(section, options);
            }
            if (layout.HorizontalRelativeFrom == "margin") horizontalMarginOrigin = margins.Left.ToPoints();
            if (!TryGetNativeGroupAxis(layout.HorizontalRelativeFrom, layout.HorizontalOffsetPoints,
                    layout.HorizontalAlignment, pageSize.Width.ToPoints(), margins.Left.ToPoints(), margins.Right.ToPoints(), layout.WidthPoints, false, out x) ||
                !TryGetNativeGroupAxis(layout.VerticalRelativeFrom, layout.VerticalOffsetPoints,
                    layout.VerticalAlignment, pageSize.Height.ToPoints(), margins.Top.ToPoints(), margins.Bottom.ToPoints(), layout.HeightPoints, true, out y)) return false;
            return x >= 0D && y >= 0D && x + layout.WidthPoints <= pageSize.Width.ToPoints() + 0.001D &&
                (paragraphRelative || y + layout.HeightPoints <= pageSize.Height.ToPoints() + 0.001D);
        }

        private static bool TryGetNativeGroupAxis(string? reference, double? offset, string? alignment,
            double pageLength, double leadingMargin, double trailingMargin, double length, bool vertical, out double position) {
            position = 0D;
            bool paragraph = vertical && reference == "paragraph";
            if (reference != "page" && reference != "margin" && !paragraph) return false;
            double origin = reference == "margin" ? leadingMargin : 0D;
            double available = reference == "margin" ? pageLength - leadingMargin - trailingMargin : pageLength;
            if (offset.HasValue) { position = origin + offset.Value; return true; }
            if (paragraph) return false;
            position = origin + (alignment == "center" ? (available - length) / 2D
                : alignment == "right" || alignment == "bottom" ? available - length : 0D);
            return alignment is "left" or "right" or "top" or "bottom" or "center";
        }

        private static bool TryAddNativeGroupChildren(OfficeDrawing scene, OfficeXmlNode group,
            double left, double top, double width, double height, WordParagraph run, OfficeXmlNode drawing,
            IReadOnlyDictionary<string, OfficeColor> theme, WordPdfExportOptions? options, int depth) {
            if (depth >= 32) throw new InvalidDataException("Word shape group nesting exceeds the PDF export limit of 32 levels.");
            OfficeXmlNode groupProperties = group.Elements().FirstOrDefault(child => child.LocalName == "grpSpPr");
            OfficeXmlNode transform = groupProperties.IsNone ? default : groupProperties.Child(A, "xfrm");
            OfficeXmlNode childExtents = transform.IsNone ? default : transform.Child(A, "chExt");
            OfficeXmlNode childOffset = transform.IsNone ? default : transform.Child(A, "chOff");
            if (transform.IsNone || transform.Int32Attr("rot") is > 0 or < 0 || ParseBoolean(transform.Attr("flipH")) == true || ParseBoolean(transform.Attr("flipV")) == true ||
                childExtents.Int64Attr("cx") is not long cw || cw <= 0 || childExtents.Int64Attr("cy") is not long ch || ch <= 0) return false;
            double cx = childOffset.Int64Attr("x") ?? 0L;
            double cy = childOffset.Int64Attr("y") ?? 0L;
            bool visible = false;
            foreach (OfficeXmlNode child in group.Elements()) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                bool nested = child.NamespaceUri == NativeWordGroupNamespace && child.LocalName == "grpSp";
                if (!nested && !child.Is(Wps, "wsp")) {
                    if (child.LocalName is "cNvPr" or "cNvGrpSpPr" or "grpSpPr" or "extLst") continue;
                    return false;
                }
                OfficeXmlNode properties = child.Elements().FirstOrDefault(item => item.LocalName == (nested ? "grpSpPr" : "spPr"));
                OfficeXmlNode childTransform = properties.IsNone ? default : properties.Elements().FirstOrDefault(item => item.LocalName == "xfrm");
                OfficeXmlNode offset = childTransform.IsNone ? default : childTransform.Child(A, "off");
                OfficeXmlNode extent = childTransform.IsNone ? default : childTransform.Child(A, "ext");
                if (extent.Int64Attr("cx") is not long ew || extent.Int64Attr("cy") is not long eh || ew <= 0 || eh <= 0) return false;
                double x = left + ((offset.Int64Attr("x") ?? 0L) - cx) * width / cw;
                double y = top + ((offset.Int64Attr("y") ?? 0L) - cy) * height / ch;
                double w = ew * width / cw, h = eh * height / ch;
                if (x < 0D || y < 0D || x + w > scene.Width + 0.001D || y + h > scene.Height + 0.001D) return false;
                if (nested) {
                    if (!TryAddNativeGroupChildren(scene, child, x, y, w, h, run, drawing, theme, options, depth + 1)) return false;
                } else {
                    if (childTransform.Is(A, "xfrm") && (childTransform.Int32Attr("rot") is > 0 or < 0 ||
                        ParseBoolean(childTransform.Attr("flipH")) == true || ParseBoolean(childTransform.Attr("flipV")) == true)) return false;
                    string? preset = properties.IsNone ? null : properties.Child(A, "prstGeom").Attr("prst");
                    OfficeShape? shape = CreateNativeDrawingPresetShape(preset, w, h);
                    if (shape == null || child.FirstDescendantSkippingExtLst(Wps, "txbx").Exists) return false;
                    WordShape wordShape = WordConverterNodes.CreateGroupChildShape(run, drawing, child);
                    ApplyNativeShapeStyle(shape, wordShape, theme);
                    scene.AddShape(shape, x, y);
                }
                visible = true;
            }
            return visible;
        }
    }
}
