using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly SvgFormRegistry svgForms;

        private bool TryDrawSvgImageAt(OfficeDrawingImage image, double originX, double originTopY, long? logicalOrder = null,
            OfficeDrawing? parsed = null) {
            bool isSvg = parsed != null || OfficeImageInfo.FromMimeType(image.ContentType) == OfficeImageFormat.Svg ||
                (OfficeImageReader.TryIdentifyByContent(image.EncodedBytes, null, out OfficeImageInfo info) && info.Format == OfficeImageFormat.Svg);
            if (!isSvg) return false;
            OfficeDrawing vector = parsed ?? ReadSvgImageForPdf(image.EncodedBytes);
            if (!CanDrawSvgSceneAsVector(vector)) {
                currentOpts.AddLayoutDiagnostic("svg-image-raster-fallback", "Drawing image",
                    "SVG image paint not supported by the PDF vector writer was rasterized.", PdfLayoutDiagnosticKind.SimplifiedContent);
                byte[] png = RasterizeSvgImageForPdf(vector, image.Projection.Width, image.Projection.Height);
                DrawDrawingImageAt(OfficeDrawingImage.CreateWithInterpolation(png, "image/png", image.Projection, image.Interpolate,
                    image.AlternativeText, image.Opacity), originX, originTopY, logicalOrder);
                return true;
            }

            SvgFormXObject? form = _drawingIsOpacityData ? null : GetOrCreateSvgForm(vector);
            OfficeImageProjection projection = image.Projection;
            bool markedContent;
            AppendDrawingMarkedContentBegin(new PdfDrawingStyle { AlternativeText = image.AlternativeText }, out markedContent, logicalOrder: logicalOrder);
            bool previousWrappers = _suppressCanvasAccessibilityWrappers;
            bool previousLinks = _suppressDrawingLinks;
            _suppressCanvasAccessibilityWrappers = true;
            _suppressDrawingLinks = true;
            try {
                RenderEffectGroup(ToTopLeftPageTransform(projection.CreateFrameTransform().CreateDestinationTransform(), originX, originTopY), image.Opacity, () => {
                    var content = new ContentStreamBuilder(sb).SaveState();
                    AppendClipPath(sb, OfficeClipPath.Rectangle(projection.Width, projection.Height),
                        originX + projection.X, originTopY - projection.Y - projection.Height, projection.Height);
                    double width = projection.Width / projection.SourceWidth;
                    double height = projection.Height / projection.SourceHeight;
                    OfficeTransform placement = OfficeTransform.Scale(width / vector.Width, height / vector.Height)
                        .Then(OfficeTransform.Translate(projection.X - projection.SourceLeft * width, projection.Y - projection.SourceTop * height));
                    RenderEffectGroup(ToTopLeftPageTransform(placement, originX, originTopY), 1D, () => {
                        if (form != null) AppendSvgFormDraw(form, originX, originTopY - vector.Height);
                        else DrawDrawingElements(vector, originX, originTopY);
                    });
                    content.RestoreState();
                });
            } finally {
                _suppressCanvasAccessibilityWrappers = previousWrappers;
                _suppressDrawingLinks = previousLinks;
                AppendDrawingMarkedContentEnd(markedContent);
            }
            return true;
        }

        private OfficeDrawing ReadSvgImageForPdf(byte[] bytes) {
            return PdfSvgImageReader.Read(bytes, remember: true, cancellationToken);
        }

        private byte[] RasterizeSvgImageForPdf(OfficeDrawing vector, double targetWidth, double targetHeight) {
            double desiredScale = Math.Max(1D, Math.Max(targetWidth / vector.Width, targetHeight / vector.Height));
            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(vector.Width, vector.Height, OfficeImageExportFormat.Png,
                new OfficeImageExportOptions { Scale = desiredScale, MaximumRasterPixels = 16_000_000L,
                    RasterOverflowBehavior = OfficeRasterOverflowBehavior.ReduceScale });
            return OfficeDrawingRasterRenderer.ToPng(vector, new OfficeDrawingRasterRenderOptions {
                Scale = plan.Limit.Scale, Background = OfficeColor.Transparent,
                CancellationToken = cancellationToken, MaximumRasterPixels = 16_000_000L,
                ThrowOnImageDecodeFailure = true, LimitRasterPaintWork = true
            });
        }

        /// <summary>Draws a shared SVG Form with its bottom-left drawing corner at the given PDF point.</summary>
        private void AppendSvgFormDraw(SvgFormXObject form, double x, double bottomY) {
            if (!currentPage!.SvgForms.Contains(form)) currentPage.SvgForms.Add(form);
            new ContentStreamBuilder(sb)
                .SaveState()
                .TransformMatrix(OfficeTransform.Translate(x, bottomY))
                .XObject(form.ResourceName)
                .RestoreState();
            pageDirty = true;
        }

        /// <summary>Returns the shared Form that draws a parsed SVG drawing, or <c>null</c> when it must be drawn inline.</summary>
        private SvgFormXObject? GetOrCreateSvgForm(OfficeDrawing vector) {
            if (svgForms.TryGet(vector, out SvgFormXObject? known)) return known;
            if (!IsSvgFormCandidate(vector) || vector.Width <= 0D || vector.Height <= 0D) {
                return svgForms.Register(vector, vector.Width, vector.Height, null,
                    Array.Empty<PageGraphicsState>(), Array.Empty<PageShading>());
            }

            EnsurePage();
            LayoutResult.Page page = currentPage!;
            // The Form names its own paint resources, so capture against empty resource lists and restore the page's.
            PageGraphicsState[] pageStates = page.GraphicsStates.ToArray();
            PageShading[] pageShadings = page.Shadings.ToArray();
            page.GraphicsStates.Clear();
            page.Shadings.Clear();
            var objects = new SvgCaptureCounts(page);
            string? content = null;
            PageGraphicsState[] formStates;
            PageShading[] formShadings;
            try {
                string captured = CaptureSvgContent(vector, opacityData: false, OfficeTransform.Identity);
                formStates = page.GraphicsStates.ToArray();
                formShadings = page.Shadings.ToArray();
                if (objects.Unchanged(page) && IsSelfContainedSvgFormContent(captured)) content = captured;
            } finally {
                page.GraphicsStates.Clear();
                page.GraphicsStates.AddRange(pageStates);
                page.Shadings.Clear();
                page.Shadings.AddRange(pageShadings);
                objects.Restore(page);
            }
            return svgForms.Register(vector, vector.Width, vector.Height, content, formStates, formShadings);
        }

        /// <summary>
        /// Writes a drawing's paint to a detached buffer in drawing units, y-up with its bottom-left at the origin.
        /// Accessibility wrappers, links and source structure stay with the enclosing image placement.
        /// </summary>
        private string CaptureSvgContent(OfficeDrawing vector, bool opacityData, OfficeTransform drawingToPage) {
            StringBuilder pageContent = sb;
            bool previousWrappers = _suppressCanvasAccessibilityWrappers;
            bool previousLinks = _suppressDrawingLinks;
            bool previousOpacityData = _drawingIsOpacityData;
            bool previousDirty = pageDirty;
            var previousSourceStructure = _drawingSourceStructure;
            OfficeTransform previousEffectToPage = _canvasEffectToPage;
            sb = new StringBuilder();
            _suppressCanvasAccessibilityWrappers = true;
            _suppressDrawingLinks = true;
            _drawingIsOpacityData = opacityData;
            _drawingSourceStructure = null;
            _canvasEffectToPage = drawingToPage;
            try {
                DrawDrawingElements(vector, 0D, vector.Height);
                return sb.ToString();
            } finally {
                sb = pageContent;
                _suppressCanvasAccessibilityWrappers = previousWrappers;
                _suppressDrawingLinks = previousLinks;
                _drawingIsOpacityData = previousOpacityData;
                _drawingSourceStructure = previousSourceStructure;
                _canvasEffectToPage = previousEffectToPage;
                pageDirty = previousDirty;
            }
        }

        /// <summary>
        /// Gives each SVG image placed on the finished page its vector paint: a shared Form when the drawing
        /// stands alone, otherwise its own content drawn inside the image placement.
        /// </summary>
        private void BindSvgPageImages() {
            LayoutResult.Page page = currentPage!;
            int count = page.Images.Count;
            for (int index = 0; index < count; index++) {
                PageImage image = page.Images[index];
                if (image.Info.Format != OfficeImageFormat.Svg || image.SvgForm != null || image.SvgContent != null) continue;
                cancellationToken.ThrowIfCancellationRequested();
                OfficeDrawing vector = image.SvgDrawing ?? ReadSvgImageForPdf(image.Data);
                if (!CanDrawSvgSceneAsVector(vector) || vector.Width <= 0D || vector.Height <= 0D) {
                    RasterizeSvgPageImage(image, vector, "SVG image paint not supported by the PDF vector writer was rasterized.");
                    continue;
                }
                if (UsesEmbeddedSvgFonts(vector)) {
                    RasterizeSvgPageImage(image, vector, "SVG image text set in fonts embedded in the SVG was rasterized at its placed size.");
                    continue;
                }

                image.SvgDrawing = vector;
                if (!image.IsOpacityData) {
                    image.SvgForm = GetOrCreateSvgForm(vector);
                    if (image.SvgForm != null) continue;
                }

                OfficeTransform drawingToPage = OfficeTransform.Scale(1D / vector.Width, 1D / vector.Height)
                    .Then(new OfficeImageProjection(new OfficeImagePlacement(image.X, image.Y, image.W, image.H),
                        rotationDegrees: image.RotationAngle, rotationCenterX: image.RotationCenterX, rotationCenterY: image.RotationCenterY,
                        flipHorizontal: image.HorizontalFlip, flipVertical: image.VerticalFlip).CreateUnitSquareTransform());
                if (image.IsForeground) drawingToPage = drawingToPage.Then(image.ForegroundTransform);
                int images = page.Images.Count, effects = page.EffectGroups.Count;
                string content = CaptureSvgContent(vector, image.IsOpacityData, drawingToPage);
                if (page.Images.Count == images && page.EffectGroups.Count == effects && !content.Contains(InlinePlaceholderMarker)) {
                    image.SvgContent = content;
                    continue;
                }

                // Nested images and transparency groups are page objects drawn where the page places them;
                // an image's own content cannot carry them, so such a drawing keeps the raster route.
                page.Images.RemoveRange(images, page.Images.Count - images);
                page.EffectGroups.RemoveRange(effects, page.EffectGroups.Count - effects);
                RasterizeSvgPageImage(image, vector, "SVG image content with nested images or transparency groups was rasterized at its placed size.");
            }
        }

        private void RasterizeSvgPageImage(PageImage image, OfficeDrawing vector, string message) {
            currentOpts.AddLayoutDiagnostic("svg-image-raster-fallback", "Image", message, PdfLayoutDiagnosticKind.SimplifiedContent);
            PdfDocument.PreparedImage raster = PdfDocument.PrepareImageBytes(RasterizeSvgImageForPdf(vector, image.W, image.H), cancellationToken);
            image.Data = raster.Data;
            image.Info = raster.Info;
            image.PreparedStream = raster.PreparedStream;
            image.SvgDrawing = null;
        }

        /// <summary>
        /// Binds an SVG image added while the page is written, such as a header or footer image. Its page
        /// resources are already fixed, so it uses a shared Form or, failing that, a raster at its placed size.
        /// </summary>
        internal void BindWrittenSvgImage(LayoutResult.Page page, PdfCreateOptions pageOptions, PageImage image) {
            currentPage = page;
            currentOpts = pageOptions;
            try {
                OfficeDrawing vector = image.SvgDrawing ?? ReadSvgImageForPdf(image.Data);
                SvgFormXObject? form = null;
                if (!image.IsOpacityData && CanDrawSvgSceneAsVector(vector) && vector.Width > 0D && vector.Height > 0D) {
                    form = GetOrCreateSvgForm(vector);
                }
                if (form != null) {
                    image.SvgDrawing = vector;
                    image.SvgForm = form;
                    return;
                }
                RasterizeSvgPageImage(image, vector, "An SVG image in running page content was rasterized at its placed size.");
            } finally {
                currentPage = null;
            }
        }

        // The PDF text writer resolves fonts from the document, not from faces an SVG embeds with @font-face;
        // text set in those faces keeps the raster route so it keeps its appearance.
        private static bool UsesEmbeddedSvgFonts(OfficeDrawing drawing) => HasScopedFonts(drawing) && ContainsSvgText(drawing);

        private static bool HasScopedFonts(OfficeDrawing drawing) {
            if (drawing.HasScopedFonts) return true;
            foreach (OfficeDrawingElement element in drawing.Elements) {
                if (element is OfficeDrawingGroup group && HasScopedFonts(group.InnerDrawing) ||
                    element is OfficeDrawingEffectGroup effect && HasScopedFonts(effect.InnerDrawing) ||
                    element is OfficeDrawingTilingPattern pattern && HasScopedFonts(pattern.InnerTile)) return true;
            }
            return false;
        }

        private static bool ContainsSvgText(OfficeDrawing drawing) {
            foreach (OfficeDrawingElement element in drawing.Elements) {
                if (element is OfficeDrawingText or OfficeDrawingRichText ||
                    element is OfficeDrawingGroup group && ContainsSvgText(group.InnerDrawing) ||
                    element is OfficeDrawingEffectGroup effect && ContainsSvgText(effect.InnerDrawing) ||
                    element is OfficeDrawingTilingPattern pattern && ContainsSvgText(pattern.InnerTile)) return true;
            }
            return false;
        }

        private static bool CanDrawSvgSceneAsVector(OfficeDrawing drawing) {
            foreach (OfficeDrawingElement element in drawing.Elements) {
                if (element is OfficeDrawingGroup group) {
                    if (!CanDrawSvgSceneAsVector(group.InnerDrawing)) return false;
                } else if (element is OfficeDrawingEffectGroup effect) {
                    if (effect.SoftMask is { } mask && (mask.Mode != OfficeSoftMaskMode.Alpha || mask.BackdropColor.A != 0 || !CanDrawSvgSceneAsVector(mask.InnerDrawing))) return false;
                    if (!CanDrawSvgSceneAsVector(effect.InnerDrawing)) return false;
                } else if (element is OfficeDrawingTilingPattern pattern) {
                    if (!CanDrawSvgSceneAsVector(pattern.InnerTile)) return false;
                } else if (element is not (OfficeDrawingShape or OfficeDrawingImage or OfficeDrawingText or OfficeDrawingRichText or OfficeDrawingLink)) return false;
            }
            return true;
        }

        /// <summary>Page objects a Form capture must leave untouched; anything it adds is removed again.</summary>
        private readonly struct SvgCaptureCounts {
            private readonly int _images, _effects, _annotations, _textAnnotations, _freeText, _highlights, _formFields,
                _destinations, _structure, _bookmarks, _drawings, _fonts, _namedFonts, _markedContentId, _imageTokenId;

            internal SvgCaptureCounts(LayoutResult.Page page) {
                _images = page.Images.Count;
                _effects = page.EffectGroups.Count;
                _annotations = page.Annotations.Count;
                _textAnnotations = page.TextAnnotations.Count;
                _freeText = page.FreeTextAnnotations.Count;
                _highlights = page.HighlightAnnotations.Count;
                _formFields = page.FormFields.Count;
                _destinations = page.NamedDestinations.Count;
                _structure = page.StructElements.Count;
                _bookmarks = page.Bookmarks.Count;
                _drawings = page.Drawings.Count;
                _fonts = page.UsedFonts.Count;
                _namedFonts = page.UsedNamedFonts.Count;
                _markedContentId = page.NextMarkedContentId;
                _imageTokenId = page.NextInlineImageTokenId;
            }

            internal bool Unchanged(LayoutResult.Page page) =>
                page.Images.Count == _images && page.EffectGroups.Count == _effects && page.Annotations.Count == _annotations &&
                page.TextAnnotations.Count == _textAnnotations && page.FreeTextAnnotations.Count == _freeText &&
                page.HighlightAnnotations.Count == _highlights && page.FormFields.Count == _formFields &&
                page.NamedDestinations.Count == _destinations && page.StructElements.Count == _structure &&
                page.Bookmarks.Count == _bookmarks && page.Drawings.Count == _drawings && page.UsedFonts.Count == _fonts &&
                page.UsedNamedFonts.Count == _namedFonts && page.NextMarkedContentId == _markedContentId &&
                page.NextInlineImageTokenId == _imageTokenId;

            internal void Restore(LayoutResult.Page page) {
                Truncate(page.Images, _images);
                Truncate(page.EffectGroups, _effects);
                Truncate(page.Annotations, _annotations);
                Truncate(page.TextAnnotations, _textAnnotations);
                Truncate(page.FreeTextAnnotations, _freeText);
                Truncate(page.HighlightAnnotations, _highlights);
                Truncate(page.FormFields, _formFields);
                Truncate(page.NamedDestinations, _destinations);
                Truncate(page.StructElements, _structure);
                Truncate(page.Bookmarks, _bookmarks);
                Truncate(page.Drawings, _drawings);
                page.NextMarkedContentId = _markedContentId;
                page.NextInlineImageTokenId = _imageTokenId;
            }

            private static void Truncate<T>(List<T> items, int count) {
                if (items.Count > count) items.RemoveRange(count, items.Count - count);
            }
        }
    }
}
