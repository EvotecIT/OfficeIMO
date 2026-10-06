using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const string NativeWordGroupNamespace = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

        private static bool HasNativeParagraphShapeGroups(IReadOnlyList<WordParagraph> runs) =>
            runs.Any(run => run.EnumerateEffectiveRunContent().SelectMany(child => child.Descendants().Prepend(child))
                .Any(child => child is V.Group || child.NamespaceUri == NativeWordGroupNamespace && child.LocalName == "wgp"));

        /// <summary>Projects a group as one object, retaining its child coordinate system and paragraph anchor.</summary>
        private static bool RenderNativeParagraphShapeGroups(INativePdfFlow pdf, WordParagraph paragraph,
            IReadOnlyList<WordParagraph> runs, PdfCore.PdfAlign align, WordToPdfOptions? options, PdfCore.PdfParagraphStyle style, ref int imageCount) {
            bool renderedFlowObject = false;
            foreach (WordParagraph run in runs) {
                foreach (W.Drawing drawing in run.EnumerateEffectiveRunContent().OfType<W.Drawing>()) {
                    OpenXmlElement? group = drawing.Descendants<A.GraphicData>().FirstOrDefault()?.ChildElements
                        .FirstOrDefault(child => child.NamespaceUri == NativeWordGroupNamespace && child.LocalName == "wgp");
                    if (group == null || !WordDrawingLayoutReader.TryRead(drawing, out WordDrawingLayoutSnapshot layout)) continue;
                    options?.CancellationToken.ThrowIfCancellationRequested();
                    if (layout.WidthPoints <= 0D || layout.HeightPoints <= 0D) {
                        WarnNativeGroup(options, "NativeShapeGroupUnsupported", "The shape group has no renderable extent.");
                        continue;
                    }

                    var scene = new OfficeDrawing(layout.WidthPoints, layout.HeightPoints);
                    var theme = GetNativeDrawingThemeColors(group.Ancestors<OpenXmlPartRootElement>().FirstOrDefault()?.OpenXmlPart);
                    bool native = TryAddNativeGroupChildren(scene, group, 0D, 0D, scene.Width, scene.Height, run, drawing, theme, options, 0);
                    if (layout.Placement == WordDrawingPlacementKind.Inline) {
                        if (native) {
                            pdf.Drawing(scene, align, spacingAfter: 0D);
                            renderedFlowObject = true;
                        }
                        else WarnNativeGroup(options, "NativeShapeGroupUnsupported", "The inline shape group contains unsupported geometry or transforms.");
                        continue;
                    }

                    if (!TryGetNativeGroupPosition(pdf, paragraph, layout, options, out double x, out double y,
                            out bool paragraphRelative, out double? horizontalMarginOrigin)) {
                        if (native) {
                            pdf.Drawing(scene, align, spacingAfter: 0D);
                            renderedFlowObject = true;
                            WarnNativeGroup(options, "NativeShapeGroupFlowed", "The shape group's anchor is outside the fixed-placement contract; it was placed in document flow.");
                        } else WarnNativeGroup(options, "NativeShapeGroupUnsupported", "The shape group's geometry and anchor could not be mapped.");
                        continue;
                    }

                    var canvas = new PdfCore.PdfPageCanvas();
                    if (native) canvas.Drawing(scene, x, y, scene.Width, scene.Height);
                    else {
                        V.Group? fallback = drawing.Ancestors<AlternateContentChoice>().FirstOrDefault()?.Parent?
                            .GetFirstChild<AlternateContentFallback>()?.Descendants<V.Group>().FirstOrDefault();
                        if (fallback == null) {
                            WarnNativeGroup(options, "NativeShapeGroupUnsupported", "The shape group contains unsupported geometry or transforms and has no VML fallback.");
                            continue;
                        }
                        CountNativeVmlGroupImages(fallback, options, ref imageCount);
                        (double cw, double ch) = GetNativeVmlCoordSize(fallback, scene.Width, scene.Height);
                        (double cx, double cy) = GetNativeVmlCoordOrigin(fallback);
                        var frame = new NativeVmlFrame(0D, 0D, scene.Width, scene.Height, cw, ch, cx, cy);
                        bool rendered = false;
                        canvas.Effect(OfficeTransform.Translate(x, y), 1D, target =>
                            rendered = RenderNativeVmlCoverChildren(target, paragraph._document, fallback.ChildElements,
                                frame, pdf.PageSize.Width, pdf.PageSize.Height));
                        if (!rendered) {
                            WarnNativeGroup(options, "NativeShapeGroupUnsupported", "The shape group's VML fallback produced no visible content.");
                            continue;
                        }
                        WarnNativeGroup(options, "NativeShapeGroupVmlFallback", "The shape group was rendered through its VML fallback because its DrawingML geometry or transforms are unsupported.");
                    }
                    AddNativeGroupCanvas(style, canvas, paragraphRelative, layout.RelativeHeight, horizontalMarginOrigin);
                }
                foreach (V.Group legacy in run.EnumerateEffectiveRunContent().SelectMany(child => child.Descendants().Prepend(child))
                    .OfType<V.Group>().Where(group => !group.Ancestors<V.Group>().Any())) {
                    RenderNativeLegacyBodyGroup(pdf, paragraph, legacy, options, style, ref imageCount);
                }
            }
            return renderedFlowObject;
        }

        private static void AddNativeGroupCanvas(PdfCore.PdfParagraphStyle style, PdfCore.PdfPageCanvas canvas,
            bool paragraphRelative, long zOrder, double? horizontalMarginOrigin = null) {
            var combined = new PdfCore.PdfPageCanvas();
            if (style.AnchoredCanvas != null) combined.AddItems(style.AnchoredCanvas.Items);
            IReadOnlyList<PdfCore.PdfCanvasItem> items = canvas.Items;
            if (horizontalMarginOrigin.HasValue)
                items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasMarginAnchorItem(items, horizontalMarginOrigin.Value) };
            if (paragraphRelative)
                items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasParagraphAnchorItem(items) };
            combined.AddItems(new[] { new PdfCore.PdfCanvasBehindTextItem(items, zOrder) });
            style.AnchoredCanvas = new PdfCore.PdfCanvasBlock(combined.Items);
        }

        private static void RenderNativeLegacyBodyGroup(INativePdfFlow pdf, WordParagraph paragraph, V.Group group,
            WordToPdfOptions? options, PdfCore.PdfParagraphStyle style, ref int imageCount) {
            Dictionary<string, string> vmlStyle = ParseNativeVmlStyle(group.Style?.Value);
            if (IsNativeVmlHidden(group)) return;
            long zIndex = 0;
            bool supportedLayer = vmlStyle.TryGetValue("z-index", out string? z) &&
                long.TryParse(z, NumberStyles.Integer, CultureInfo.InvariantCulture, out zIndex) && zIndex < 0;
            bool pageX = vmlStyle.TryGetValue("mso-position-horizontal-relative", out string? horizontal) && horizontal == "page";
            bool pageY = vmlStyle.TryGetValue("mso-position-vertical-relative", out string? vertical) && vertical == "page";
            if (!supportedLayer || !pageX || (!pageY && vertical != null && vertical != "text")) {
                WarnNativeGroup(options, "NativeShapeGroupUnsupported", "The legacy shape group's anchor is outside the non-wrapping behind-text placement contract.");
                return;
            }
            CountNativeVmlGroupImages(group, options, ref imageCount);
            var frame = new NativeVmlFrame(0D, 0D, pdf.PageSize.Width, pdf.PageSize.Height,
                pdf.PageSize.Width, pdf.PageSize.Height, 0D, 0D);
            var canvas = new PdfCore.PdfPageCanvas();
            if (RenderNativeVmlGroup(canvas, paragraph._document, group, frame, pdf.PageSize.Width, pdf.PageSize.Height))
                AddNativeGroupCanvas(style, canvas, paragraphRelative: !pageY, zOrder: zIndex);
            else WarnNativeGroup(options, "NativeShapeGroupUnsupported", "The legacy shape group produced no visible content.");
        }

        private static void CountNativeVmlGroupImages(V.Group group, WordToPdfOptions? options, ref int imageCount) {
            int imageLimit = options?.MaxImagesPerParagraph ?? 1_000;
            if (imageLimit <= 0) throw new ArgumentOutOfRangeException(nameof(WordToPdfOptions.MaxImagesPerParagraph));
            foreach (V.ImageData image in group.Descendants<V.ImageData>()) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (image.Ancestors<V.Shape>().Any(IsNativeVmlHidden) || image.Ancestors<V.Group>().Any(IsNativeVmlHidden)) continue;
                if (++imageCount > imageLimit)
                    throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
            }
        }

        private static void WarnNativeGroup(WordToPdfOptions? options, string code, string message) {
            if (options != null) AddNativeExportWarning(options, code, "body paragraph shape group", message);
        }

        private static bool TryGetNativeGroupPosition(INativePdfFlow pdf, WordParagraph paragraph,
            WordDrawingLayoutSnapshot layout, WordToPdfOptions? options, out double x, out double y,
            out bool paragraphRelative, out double? horizontalMarginOrigin) {
            x = y = 0D;
            horizontalMarginOrigin = null;
            paragraphRelative = layout.VerticalRelativeFrom == "paragraph";
            if (layout.Wrap != WordDrawingWrapKind.None || !layout.BehindDocument) return false;
            PdfCore.PageMargins margins = PdfCore.PageMargins.Uniform(0D);
            if (layout.HorizontalRelativeFrom == "margin" || layout.VerticalRelativeFrom == "margin") {
                if (paragraph.Parent is not WordSection section) return false;
                margins = GetNativeMargins(section, options);
            }
            if (layout.HorizontalRelativeFrom == "margin") horizontalMarginOrigin = margins.Left;
            if (!TryGetNativeGroupAxis(layout.HorizontalRelativeFrom, layout.HorizontalOffsetPoints,
                    layout.HorizontalAlignment, pdf.PageSize.Width, margins.Left, margins.Right, layout.WidthPoints, false, out x) ||
                !TryGetNativeGroupAxis(layout.VerticalRelativeFrom, layout.VerticalOffsetPoints,
                    layout.VerticalAlignment, pdf.PageSize.Height, margins.Top, margins.Bottom, layout.HeightPoints, true, out y)) return false;
            return x >= 0D && y >= 0D && x + layout.WidthPoints <= pdf.PageSize.Width + 0.001D &&
                (paragraphRelative || y + layout.HeightPoints <= pdf.PageSize.Height + 0.001D);
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

        private static bool TryAddNativeGroupChildren(OfficeDrawing scene, OpenXmlElement group,
            double left, double top, double width, double height, WordParagraph run, W.Drawing drawing,
            IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> theme, WordToPdfOptions? options, int depth) {
            if (depth >= 32) throw new InvalidDataException("Word shape group nesting exceeds the PDF export limit of 32 levels.");
            A.TransformGroup? transform = group.ChildElements.FirstOrDefault(child => child.LocalName == "grpSpPr")?.GetFirstChild<A.TransformGroup>();
            if (transform == null || transform.Rotation?.Value is > 0 or < 0 || transform.HorizontalFlip?.Value == true || transform.VerticalFlip?.Value == true ||
                transform.ChildExtents?.Cx?.Value is not long cw || cw <= 0 || transform.ChildExtents.Cy?.Value is not long ch || ch <= 0) return false;
            double cx = transform.ChildOffset?.X?.Value ?? 0L;
            double cy = transform.ChildOffset?.Y?.Value ?? 0L;
            bool visible = false;
            foreach (OpenXmlElement child in group.ChildElements) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                bool nested = child.NamespaceUri == NativeWordGroupNamespace && child.LocalName == "grpSp";
                if (!nested && child is not Wps.WordprocessingShape) {
                    if (child.LocalName is "cNvPr" or "cNvGrpSpPr" or "grpSpPr" or "extLst") continue;
                    return false;
                }
                OpenXmlElement? properties = child.ChildElements.FirstOrDefault(item => item.LocalName == (nested ? "grpSpPr" : "spPr"));
                OpenXmlElement? childTransform = properties?.ChildElements.FirstOrDefault(item => item.LocalName == "xfrm");
                A.Offset? offset = childTransform?.GetFirstChild<A.Offset>();
                A.Extents? extent = childTransform?.GetFirstChild<A.Extents>();
                if (extent?.Cx?.Value is not long ew || extent.Cy?.Value is not long eh || ew <= 0 || eh <= 0) return false;
                double x = left + ((offset?.X?.Value ?? 0L) - cx) * width / cw;
                double y = top + ((offset?.Y?.Value ?? 0L) - cy) * height / ch;
                double w = ew * width / cw, h = eh * height / ch;
                if (x < 0D || y < 0D || x + w > scene.Width + 0.001D || y + h > scene.Height + 0.001D) return false;
                if (nested) {
                    if (!TryAddNativeGroupChildren(scene, child, x, y, w, h, run, drawing, theme, options, depth + 1)) return false;
                } else {
                    if (childTransform is A.Transform2D t && (t.Rotation?.Value is > 0 or < 0 || t.HorizontalFlip?.Value == true || t.VerticalFlip?.Value == true)) return false;
                    string? preset = properties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText;
                    OfficeShape? shape = CreateNativeDrawingPresetShape(preset, w, h);
                    if (shape == null || child.Descendants<Wps.TextBoxInfo2>().Any()) return false;
                    var wordShape = new WordShape(run._document, run._paragraph!, run._run!, drawing) { _wpsShape = (Wps.WordprocessingShape)child };
                    ApplyNativeShapeStyle(shape, wordShape, theme);
                    scene.AddShape(shape, x, y);
                }
                visible = true;
            }
            return visible;
        }
    }
}
