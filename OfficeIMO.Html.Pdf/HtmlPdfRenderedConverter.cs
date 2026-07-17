using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static class HtmlPdfRenderedConverter {
    private const double PointsPerCssPixel = 72D / HtmlRenderOptions.CssPixelsPerInch;

    internal static HtmlPdfRenderResult Convert(HtmlConversionDocument document, HtmlPdfSaveOptions options) {
        HtmlRenderOptions renderOptions = ResolveRenderOptions(options);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, renderOptions);
        return CreatePdf(rendered, options, CancellationToken.None);
    }

    internal static async Task<HtmlPdfRenderResult> ConvertAsync(HtmlConversionDocument document, HtmlPdfSaveOptions options, CancellationToken cancellationToken) {
        HtmlRenderOptions renderOptions = ResolveRenderOptions(options);
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(document, renderOptions, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return CreatePdf(rendered, options, cancellationToken);
    }

    private static HtmlRenderOptions ResolveRenderOptions(HtmlPdfSaveOptions options) {
        HtmlRenderOptions renderOptions = options.ClonePdf();
        renderOptions.Mode = HtmlRenderMode.Paged;
        renderOptions.UrlPolicy.AllowDataUrls = options.ResourcePolicy.AllowDataUris;
        HtmlRenderResourceResolver? embeddedPackageResolver = options.EmbeddedPackageResourceResolver;
        renderOptions.UrlPolicy.DisallowFileUrls = !options.ResourcePolicy.AllowLocalFileAccess &&
            !(embeddedPackageResolver != null && options.ResourcePolicy.AllowEmbeddedPackageResources);
        HtmlRenderResourceResolver? hostResolver = renderOptions.ResourceResolver;
        if (embeddedPackageResolver != null || hostResolver != null) {
            renderOptions.ResourceResolver = async (request, cancellationToken) => {
                if (embeddedPackageResolver != null && options.ResourcePolicy.AllowEmbeddedPackageResources) {
                    HtmlResolvedResource? embedded = await embeddedPackageResolver(request, cancellationToken).ConfigureAwait(false);
                    if (embedded != null) return embedded;
                }

                if (hostResolver == null) return null;
                bool hostResourceAllowed = request.Uri.IsFile
                    ? options.ResourcePolicy.AllowLocalFileAccess
                    : options.ResourcePolicy.AllowRemoteResourceResolution;
                return hostResourceAllowed
                    ? await hostResolver(request, cancellationToken).ConfigureAwait(false)
                    : null;
            };
        }
        return renderOptions;
    }

    private static HtmlPdfRenderResult CreatePdf(HtmlRenderDocument rendered, HtmlPdfSaveOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlDiagnosticReport diagnostics = rendered.DiagnosticReport.Clone();

        var conversionReport = new PdfCore.PdfConversionReport();
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create()
            .TaggedPdfCatalogMarkers();
        pdf.Options.ReportDiagnosticsTo(conversionReport, "OfficeIMO.Html.Pdf");
        if (rendered.Metadata.Title != null) pdf.Meta(title: rendered.Metadata.Title);
        if (rendered.Metadata.Language != null) pdf.Language(rendered.Metadata.Language);
        if (rendered.Metadata.Title != null || rendered.Metadata.Direction == HtmlRenderTextDirection.RightToLeft) {
            pdf.ViewerPreferences(preferences => {
                if (rendered.Metadata.Title != null) preferences.DisplayDocTitle = true;
                if (rendered.Metadata.Direction == HtmlRenderTextDirection.RightToLeft) preferences.Direction = PdfCore.PdfViewerDirection.RightToLeft;
            });
        }
        if (options.FontFamily != null) {
            pdf.UseFontFamily(options.FontFamily);
        }

        var reservedFontSlots = new HashSet<PdfCore.PdfStandardFont>();
        if (options.FontFamily != null) reservedFontSlots.Add(PdfCore.PdfStandardFont.Helvetica);
        var activeWebFontFamilies = new HashSet<string>(
            rendered.Fonts.Faces.Select(face => face.FamilyName),
            StringComparer.OrdinalIgnoreCase);
        PdfCore.PdfTextFallbackFeatures activeTextFallbacks = ResolveTextFallbackFeatures(rendered, options.TextFallbacks);
        if (activeTextFallbacks != PdfCore.PdfTextFallbackFeatures.None && options.ResourcePolicy.AllowSystemFontEmbedding) {
            RegisterUsedSystemFontFamilies(pdf, rendered, activeWebFontFamilies, reservedFontSlots);
        }
        ReserveUsedStandardFontSlots(rendered, activeWebFontFamilies, reservedFontSlots);
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts = RegisterWebFonts(
            pdf,
            rendered,
            diagnostics,
            reservedFontSlots,
            cancellationToken);
        foreach (PdfCore.PdfStandardFont slot in webFonts.Values) {
            reservedFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(slot));
        }
        if (activeTextFallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            pdf.Options.UseTextFallbacks(activeTextFallbacks, reservedFontSlots, options.ResourcePolicy.AllowSystemFontEmbedding);
        }
        pdf.UseTextShaping(options.TextShapingMode, options.TextShapingProvider);
        ILookup<int, HtmlRenderHeading> headingsByPage = rendered.Headings.ToLookup(heading => heading.PageNumber);
        foreach (HtmlRenderPage renderedPage in rendered.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            double pageWidth = renderedPage.Width * PointsPerCssPixel;
            double pageHeight = renderedPage.Height * PointsPerCssPixel;
            pdf.Page(page => page
                .Size(pageWidth, pageHeight)
                .Margin(0D)
                .Canvas(canvas => {
                    AddPageVisuals(canvas, renderedPage, webFonts, cancellationToken);
                    AddPageOutlines(canvas, headingsByPage[renderedPage.PageNumber], cancellationToken);
                }));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlPdfRenderResult(pdf, diagnostics, conversionReport);
    }

    private static void AddPageOutlines(PdfCore.PdfPageCanvas canvas, IEnumerable<HtmlRenderHeading> headings, CancellationToken cancellationToken) {
        foreach (HtmlRenderHeading heading in headings) {
            cancellationToken.ThrowIfCancellationRequested();
            canvas.Outline(heading.Text, heading.Level, heading.Y * PointsPerCssPixel);
        }
    }

    private static void AddPageVisuals(PdfCore.PdfPageCanvas canvas, HtmlRenderPage page, IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts, CancellationToken cancellationToken) {
        foreach (HtmlRenderVisual visual in page.Scene.OrderBy(item => item.PaintOrder)) {
            cancellationToken.ThrowIfCancellationRequested();
            AddVisual(canvas, visual, webFonts, page.Width, page.Height, cancellationToken);
        }
    }

    private static void AddVisual(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderVisual visual,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        double surfaceWidth,
        double surfaceHeight,
        CancellationToken cancellationToken,
        bool textAsSpan = false) {
        cancellationToken.ThrowIfCancellationRequested();
        if (visual is HtmlRenderShape shape) {
            AddShape(canvas, shape);
        } else if (visual is HtmlRenderText text) {
            AddText(canvas, text, webFonts, textAsSpan);
        } else if (visual is HtmlRenderImage image) {
            AddImage(canvas, image);
        } else if (visual is HtmlRenderDrawing drawing) {
            AddDrawing(canvas, drawing, webFonts, cancellationToken);
        } else if (visual is HtmlRenderImagePattern imagePattern) {
            AddImagePattern(canvas, imagePattern, cancellationToken);
        } else if (visual is HtmlRenderClipGroup group) {
            AddClipGroup(canvas, group, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
        } else if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            AddPathClipGroup(canvas, pathClipGroup, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
        } else if (visual is HtmlRenderEffectGroup effectGroup) {
            AddEffectGroup(canvas, effectGroup, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
        } else if (visual is HtmlRenderSemanticGroup semanticGroup) {
            AddSemanticGroup(canvas, semanticGroup, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
        } else if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            AddLogicalTextGroup(canvas, logicalTextGroup, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
        }
    }

    private static void AddLogicalTextGroup(PdfCore.PdfPageCanvas canvas, HtmlRenderLogicalTextGroup group, IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts, double surfaceWidth, double surfaceHeight, CancellationToken cancellationToken, bool textAsSpan) {
        canvas.ActualText(group.Text, nested => {
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(nested, child, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
            }
        });
    }

    private static void AddSemanticGroup(PdfCore.PdfPageCanvas canvas, HtmlRenderSemanticGroup group, IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts, double surfaceWidth, double surfaceHeight, CancellationToken cancellationToken, bool textAsSpan) {
        if (!group.Visuals.Any(ContainsPaintableVisual)) return;
        var options = new PdfCore.PdfCanvasStructureOptions {
            ColumnSpan = group.ColumnSpan,
            RowSpan = group.RowSpan,
            HeaderScope = MapTableHeaderScope(group.HeaderScope)
        };
        bool childTextAsSpan = textAsSpan || IsTextContentGroup(group.Role);
        canvas.Structure(MapSemanticGroupRole(group.Role), nested => {
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(nested, child, webFonts, surfaceWidth, surfaceHeight, cancellationToken, childTextAsSpan);
            }
        }, options);
    }

    private static bool ContainsPaintableVisual(HtmlRenderVisual visual) {
        if (visual is HtmlRenderSemanticGroup semanticGroup) return semanticGroup.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) return logicalTextGroup.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderClipGroup clipGroup) return clipGroup.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderPathClipGroup pathClipGroup) return pathClipGroup.Visuals.Any(ContainsPaintableVisual);
        if (visual is HtmlRenderEffectGroup effectGroup) return effectGroup.Visuals.Any(ContainsPaintableVisual);
        return true;
    }

    private static bool IsTextContentGroup(HtmlRenderSemanticGroupRole role) =>
        role == HtmlRenderSemanticGroupRole.Paragraph
        || role == HtmlRenderSemanticGroupRole.Heading1
        || role == HtmlRenderSemanticGroupRole.Heading2
        || role == HtmlRenderSemanticGroupRole.Heading3
        || role == HtmlRenderSemanticGroupRole.Heading4
        || role == HtmlRenderSemanticGroupRole.Heading5
        || role == HtmlRenderSemanticGroupRole.Heading6;

    private static PdfCore.PdfCanvasTableHeaderScope? MapTableHeaderScope(HtmlRenderTableHeaderScope? scope) {
        if (scope == HtmlRenderTableHeaderScope.Row) return PdfCore.PdfCanvasTableHeaderScope.Row;
        if (scope == HtmlRenderTableHeaderScope.Column) return PdfCore.PdfCanvasTableHeaderScope.Column;
        if (scope == HtmlRenderTableHeaderScope.Both) return PdfCore.PdfCanvasTableHeaderScope.Both;
        return null;
    }

    private static PdfCore.PdfCanvasStructureRole MapSemanticGroupRole(HtmlRenderSemanticGroupRole role) {
        if (role == HtmlRenderSemanticGroupRole.Section) return PdfCore.PdfCanvasStructureRole.Section;
        if (role == HtmlRenderSemanticGroupRole.Division) return PdfCore.PdfCanvasStructureRole.Division;
        if (role == HtmlRenderSemanticGroupRole.Paragraph) return PdfCore.PdfCanvasStructureRole.Paragraph;
        if (role == HtmlRenderSemanticGroupRole.Heading1) return PdfCore.PdfCanvasStructureRole.Heading1;
        if (role == HtmlRenderSemanticGroupRole.Heading2) return PdfCore.PdfCanvasStructureRole.Heading2;
        if (role == HtmlRenderSemanticGroupRole.Heading3) return PdfCore.PdfCanvasStructureRole.Heading3;
        if (role == HtmlRenderSemanticGroupRole.Heading4) return PdfCore.PdfCanvasStructureRole.Heading4;
        if (role == HtmlRenderSemanticGroupRole.Heading5) return PdfCore.PdfCanvasStructureRole.Heading5;
        if (role == HtmlRenderSemanticGroupRole.Heading6) return PdfCore.PdfCanvasStructureRole.Heading6;
        if (role == HtmlRenderSemanticGroupRole.List) return PdfCore.PdfCanvasStructureRole.List;
        if (role == HtmlRenderSemanticGroupRole.ListItem) return PdfCore.PdfCanvasStructureRole.ListItem;
        if (role == HtmlRenderSemanticGroupRole.ListLabel) return PdfCore.PdfCanvasStructureRole.ListLabel;
        if (role == HtmlRenderSemanticGroupRole.ListBody) return PdfCore.PdfCanvasStructureRole.ListBody;
        if (role == HtmlRenderSemanticGroupRole.Table) return PdfCore.PdfCanvasStructureRole.Table;
        if (role == HtmlRenderSemanticGroupRole.TableRow) return PdfCore.PdfCanvasStructureRole.TableRow;
        if (role == HtmlRenderSemanticGroupRole.TableHeaderCell) return PdfCore.PdfCanvasStructureRole.TableHeaderCell;
        if (role == HtmlRenderSemanticGroupRole.TableCell) return PdfCore.PdfCanvasStructureRole.TableCell;
        return PdfCore.PdfCanvasStructureRole.Caption;
    }

    private static void AddEffectGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderEffectGroup group,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        double surfaceWidth,
        double surfaceHeight,
        CancellationToken cancellationToken,
        bool textAsSpan) {
        OfficeTransform transform = group.Transform;
        var scaled = new OfficeTransform(
            transform.M11,
            transform.M12,
            transform.M21,
            transform.M22,
            transform.OffsetX * PointsPerCssPixel,
            transform.OffsetY * PointsPerCssPixel);
        canvas.Effect(scaled, group.Opacity, nested => {
            foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                cancellationToken.ThrowIfCancellationRequested();
                AddVisual(nested, child, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
            }
        });
    }

    private static void AddClipGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderClipGroup group,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        double surfaceWidth,
        double surfaceHeight,
        CancellationToken cancellationToken,
        bool textAsSpan) {
        double left = group.ClipHorizontal ? Math.Max(0D, group.ClipX) : 0D;
        double top = group.ClipVertical ? Math.Max(0D, group.ClipY) : 0D;
        double right = group.ClipHorizontal ? Math.Min(surfaceWidth, group.ClipX + group.ClipWidth) : surfaceWidth;
        double bottom = group.ClipVertical ? Math.Min(surfaceHeight, group.ClipY + group.ClipHeight) : surfaceHeight;
        if (right <= left + 0.0001D || bottom <= top + 0.0001D) return;
        canvas.Clip(
            left * PointsPerCssPixel,
            top * PointsPerCssPixel,
            (right - left) * PointsPerCssPixel,
            (bottom - top) * PointsPerCssPixel,
            clipped => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(clipped, child, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
                }
            });
    }

    private static void AddPathClipGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderPathClipGroup group,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        double surfaceWidth,
        double surfaceHeight,
        CancellationToken cancellationToken,
        bool textAsSpan) {
        canvas.Clip(
            group.ClipX * PointsPerCssPixel,
            group.ClipY * PointsPerCssPixel,
            group.ClipPath.Scale(PointsPerCssPixel, PointsPerCssPixel),
            clipped => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddVisual(clipped, child, webFonts, surfaceWidth, surfaceHeight, cancellationToken, textAsSpan);
                }
            });
    }

    private static void AddShape(PdfCore.PdfPageCanvas canvas, HtmlRenderShape visual) {
        var drawing = new OfficeDrawing(visual.Width, visual.Height);
        drawing.AddShape(visual.Shape.Clone(), 0D, 0D);
        canvas.Drawing(
            drawing,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            linkUri: visual.LinkUri,
            linkContents: visual.LinkUri == null ? null : visual.Source);
    }

    private static void AddText(PdfCore.PdfPageCanvas canvas, HtmlRenderText visual, IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts, bool asSpan) {
        if (visual.Text.Length == 0) return;
        string? link = string.IsNullOrWhiteSpace(visual.Text) ? null : visual.LinkUri;
        var run = new PdfCore.TextRun(
            visual.Text,
            bold: visual.Font.IsBold,
            underline: visual.Font.IsUnderline,
            color: PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            italic: visual.Font.IsItalic,
            strike: visual.Font.IsStrikethrough,
            fontSize: visual.Font.Size * PointsPerCssPixel,
            font: MapFont(visual.Font.FamilyName, webFonts),
            linkUri: link,
            linkContents: link == null ? null : visual.Text);
        canvas.Text(
            new[] { run },
            asSpan ? PdfCore.PdfCanvasTextStructureRole.Span : MapStructureRole(visual.SemanticRole),
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            MapAlignment(visual.Alignment),
            visual.Font.Size * PointsPerCssPixel,
            visual.LineHeight * PointsPerCssPixel);
    }

    private static PdfCore.PdfCanvasTextStructureRole MapStructureRole(string? semanticRole) {
        if (semanticRole == "heading-1") return PdfCore.PdfCanvasTextStructureRole.Heading1;
        if (semanticRole == "heading-2") return PdfCore.PdfCanvasTextStructureRole.Heading2;
        if (semanticRole == "heading-3") return PdfCore.PdfCanvasTextStructureRole.Heading3;
        if (semanticRole == "heading-4") return PdfCore.PdfCanvasTextStructureRole.Heading4;
        if (semanticRole == "heading-5") return PdfCore.PdfCanvasTextStructureRole.Heading5;
        if (semanticRole == "heading-6") return PdfCore.PdfCanvasTextStructureRole.Heading6;
        return semanticRole == "span" ? PdfCore.PdfCanvasTextStructureRole.Span : PdfCore.PdfCanvasTextStructureRole.Paragraph;
    }

    private static void AddImage(PdfCore.PdfPageCanvas canvas, HtmlRenderImage visual) {
        if (!TryPreparePdfImageBytes(visual.Bytes, visual.ContentType, out byte[] imageBytes)) return;
        PdfCore.PdfImageStyle? style = visual.SourceCrop.HasCrop
            ? new PdfCore.PdfImageStyle {
                SourceCrop = new PdfCore.PdfImageSourceCrop(
                    visual.SourceCrop.Left,
                    visual.SourceCrop.Top,
                    visual.SourceCrop.Right,
                    visual.SourceCrop.Bottom)
            }
            : null;
        canvas.Image(
            imageBytes,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            style,
            linkUri: visual.LinkUri,
            linkContents: visual.LinkUri == null ? null : visual.Source,
            alternativeText: visual.AlternativeText);
    }

    private static void AddDrawing(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderDrawing visual,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        CancellationToken cancellationToken) {
        OfficeDrawing source = visual.Drawing;
        double scaleX = visual.Width / source.Width;
        double scaleY = visual.Height / source.Height;
        double originX = visual.X * PointsPerCssPixel;
        double originY = visual.Y * PointsPerCssPixel;
        OfficeTransform drawingToPage = OfficeTransform.Scale(scaleX * PointsPerCssPixel, scaleY * PointsPerCssPixel)
            .Then(OfficeTransform.Translate(originX, originY));
        OfficeTransform pageToDrawing = drawingToPage.Invert();

        void AddElements(PdfCore.PdfPageCanvas target, IReadOnlyList<OfficeDrawingElement> elements) {
            var shapeBatch = new OfficeDrawing(source.Width, source.Height);
            void FlushShapes() {
                if (shapeBatch.Elements.Count == 0) return;
                cancellationToken.ThrowIfCancellationRequested();
                target.Drawing(
                    shapeBatch,
                    originX,
                    originY,
                    visual.Width * PointsPerCssPixel,
                    visual.Height * PointsPerCssPixel,
                    linkUri: visual.LinkUri,
                    linkContents: visual.LinkUri == null ? null : visual.Source);
                shapeBatch = new OfficeDrawing(source.Width, source.Height);
            }

            foreach (OfficeDrawingElement element in elements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element is OfficeDrawingShape shape) {
                    shapeBatch.AddShape(shape.Shape, shape.X, shape.Y);
                    continue;
                }
                if (element is OfficeDrawingEffectGroup effectGroup) {
                    FlushShapes();
                    OfficeTransform pageTransform = pageToDrawing
                        .Then(effectGroup.Transform)
                        .Then(drawingToPage);
                    OfficeDrawing nestedDrawing = effectGroup.Drawing;
                    target.Effect(pageTransform, effectGroup.Opacity, nested => AddElements(nested, nestedDrawing.Elements));
                    continue;
                }
                if (element is not OfficeDrawingText text || text.Text.Length == 0) continue;
                FlushShapes();
                double fontSize = text.Font.Size * scaleY * PointsPerCssPixel;
                double lineHeight = (text.LineHeight ?? text.Font.Size * 1.2D) * scaleY * PointsPerCssPixel;
                PdfCore.PdfColor? color = text.Color.HasValue ? PdfCore.PdfColor.FromOfficeColorOrNull(text.Color.Value) : null;
                var run = new PdfCore.TextRun(
                    text.Text,
                    bold: text.Font.IsBold,
                    underline: text.Font.IsUnderline,
                    color: color,
                    italic: text.Font.IsItalic,
                    strike: text.Font.IsStrikethrough,
                    fontSize: fontSize,
                    font: MapFont(text.Font.FamilyName, webFonts),
                    linkUri: visual.LinkUri,
                    linkContents: visual.LinkUri == null ? null : text.Text);
                target.Text(
                    new[] { run },
                    (visual.X + text.X * scaleX) * PointsPerCssPixel,
                    (visual.Y + text.Y * scaleY) * PointsPerCssPixel,
                    text.Width * scaleX * PointsPerCssPixel,
                    text.Height * scaleY * PointsPerCssPixel,
                    color,
                    MapAlignment(text.Alignment),
                    fontSize,
                    lineHeight);
            }
            FlushShapes();
        }

        if (string.IsNullOrWhiteSpace(visual.AlternativeText)) {
            AddElements(canvas, source.Elements);
        } else {
            canvas.Figure(visual.AlternativeText!, figure => AddElements(figure, source.Elements));
        }
    }

    private static void AddImagePattern(PdfCore.PdfPageCanvas canvas, HtmlRenderImagePattern visual, CancellationToken cancellationToken) {
        if (!TryPreparePdfImageBytes(visual.Bytes, visual.ContentType, out byte[] imageBytes)) return;
        OfficeImagePatternLayout pattern = visual.Pattern.Scale(PointsPerCssPixel);
        OfficeImagePlacement area = pattern.Area;
        PdfCore.PdfCanvasImageResource imageResource = PdfCore.PdfCanvasImageResource.Create(imageBytes);
        canvas.Clip(area.X, area.Y, area.Width, area.Height, clipped => {
            foreach (OfficeImagePlacement tile in pattern.GetTilePlacements(visual.MaximumTileCount)) {
                cancellationToken.ThrowIfCancellationRequested();
                clipped.ImageShared(imageResource, tile.X, tile.Y, tile.Width, tile.Height);
            }
        });
    }

    private static PdfCore.PdfStandardFont MapFont(string familyName, IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts) {
        foreach (string candidate in EnumerateFamilies(familyName)) {
            if (webFonts.TryGetValue(candidate, out PdfCore.PdfStandardFont embedded)) {
                return embedded;
            }
        }

        return MapStandardFont(familyName);
    }

    private static PdfCore.PdfStandardFont MapStandardFont(string familyName) {
        return PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont font)
            ? font
            : PdfCore.PdfStandardFont.Helvetica;
    }

    private static IReadOnlyDictionary<string, PdfCore.PdfStandardFont> RegisterWebFonts(
        PdfCore.PdfDocument pdf,
        HtmlRenderDocument rendered,
        HtmlDiagnosticReport? diagnostics,
        ISet<PdfCore.PdfStandardFont> reservedFontSlots,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeFontFaceCollection faces = rendered.Fonts;
        var byFamily = faces.Faces
            .GroupBy(face => face.FamilyName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        var mappings = new Dictionary<string, PdfCore.PdfStandardFont>(StringComparer.OrdinalIgnoreCase);
        if (byFamily.Count == 0) {
            return mappings;
        }

        var orderedFamilies = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string familyNames in EnumerateUsedFontFamilyLists(rendered.Pages.SelectMany(page => page.Visuals))) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (string family in EnumerateFamilies(familyNames)) {
                if (byFamily.ContainsKey(family) && seen.Add(family)) {
                    orderedFamilies.Add(family);
                }
            }
        }

        PdfCore.PdfStandardFont[] slots = {
            PdfCore.PdfStandardFont.Helvetica,
            PdfCore.PdfStandardFont.TimesRoman,
            PdfCore.PdfStandardFont.Courier
        };
        PdfCore.PdfStandardFont[] availableSlots = slots
            .Where(slot => !reservedFontSlots.Contains(PdfCore.PdfStandardFontMapper.GetFontFamily(slot)))
            .ToArray();
        for (int index = 0; index < orderedFamilies.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            string family = orderedFamilies[index];
            if (index >= availableSlots.Length) {
                diagnostics?.Add(
                    "OfficeIMO.Html.Pdf",
                    HtmlPdfDiagnosticCodes.RenderedFontFamilyLimitExceeded,
                    "The rendered PDF can embed three distinct active web-font families; an additional family used standard-font fallback.",
                    HtmlDiagnosticSeverity.Warning,
                    family,
                    "limit=" + availableSlots.Length);
                continue;
            }

            PdfCore.PdfStandardFont slot = availableSlots[index];
            RegisterFamily(pdf, slot, family, byFamily[family], cancellationToken);
            mappings[family] = slot;
        }

        return mappings;
    }

    private static void ReserveUsedStandardFontSlots(
        HtmlRenderDocument rendered,
        ISet<string> activeWebFontFamilies,
        ISet<PdfCore.PdfStandardFont> reservedFontSlots) {
        foreach (string familyNames in EnumerateUsedFontFamilyLists(rendered.Pages.SelectMany(page => page.Visuals))) {
            if (EnumerateFamilies(familyNames).Any(activeWebFontFamilies.Contains)) continue;
            reservedFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(MapStandardFont(familyNames)));
        }
    }

    private static void RegisterUsedSystemFontFamilies(
        PdfCore.PdfDocument pdf,
        HtmlRenderDocument rendered,
        ISet<string> activeWebFontFamilies,
        ISet<PdfCore.PdfStandardFont> reservedFontSlots) {
        List<HtmlRenderText> textRuns = EnumerateVisuals(rendered.Pages.SelectMany(page => page.Visuals))
            .OfType<HtmlRenderText>()
            .Where(text => !EnumerateFamilies(text.Font.FamilyName).Any(activeWebFontFamilies.Contains))
            .ToList();

        foreach (IGrouping<PdfCore.PdfStandardFont, HtmlRenderText> slotRuns in textRuns.GroupBy(
                     text => PdfCore.PdfStandardFontMapper.GetFontFamily(MapStandardFont(text.Font.FamilyName)))) {
            PdfCore.PdfStandardFont slot = slotRuns.Key;
            if (pdf.Options.HasEmbeddedStandardFontFamily(slot)) {
                reservedFontSlots.Add(slot);
                continue;
            }

            foreach (string familyName in slotRuns.SelectMany(text => EnumerateFamilies(text.Font.FamilyName)).Distinct(StringComparer.OrdinalIgnoreCase)) {
                if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfCore.PdfEmbeddedFontFamily? family) || family == null) {
                    continue;
                }

                pdf.Options.RegisterFontFamily(slot, CreateCoverageSafeFontFamily(family, slotRuns));
                reservedFontSlots.Add(slot);
                break;
            }
        }
    }

    private static PdfCore.PdfEmbeddedFontFamily CreateCoverageSafeFontFamily(
        PdfCore.PdfEmbeddedFontFamily family,
        IEnumerable<HtmlRenderText> textRuns) {
        List<HtmlRenderText> runs = textRuns.ToList();
        byte[] regular = family.Regular;
        byte[]? bold = SelectCoverageSafeFace(family.Bold, regular, runs.Where(run => run.Font.IsBold && !run.Font.IsItalic).Select(run => run.Text));
        byte[]? italic = SelectCoverageSafeFace(family.Italic, regular, runs.Where(run => !run.Font.IsBold && run.Font.IsItalic).Select(run => run.Text));
        byte[]? boldItalic = SelectCoverageSafeFace(
            family.BoldItalic ?? family.Bold ?? family.Italic,
            regular,
            runs.Where(run => run.Font.IsBold && run.Font.IsItalic).Select(run => run.Text));
        return new PdfCore.PdfEmbeddedFontFamily(family.FamilyName, regular, bold, italic, boldItalic);
    }

    private static byte[]? SelectCoverageSafeFace(byte[]? styledFace, byte[] regularFace, IEnumerable<string> requiredText) {
        if (styledFace == null) return null;
        string text = string.Concat(requiredText);
        if (text.Length == 0 || FontCoversText(styledFace, text) || !FontCoversText(regularFace, text)) return styledFace;
        return regularFace;
    }

    private static bool FontCoversText(byte[] fontData, string text) {
        var candidate = new PdfCore.PdfEmbeddedFontFallbackCandidate("HTML system font coverage", fontData);
        return PdfCore.PdfTextDiagnostics.PlanEmbeddedFontFallbackText(text, new[] { candidate }).IsFullyCovered;
    }

    private static bool TryPreparePdfImageBytes(byte[] bytes, string contentType, out byte[] pdfBytes) {
        OfficeImageFormat format = OfficeImageInfo.FromMimeType(contentType);
        string extension = OfficeImageInfo.GetDefaultExtension(format);
        if (OfficeImageReader.TryIdentify(bytes, extension, out OfficeImageInfo identified)) {
            format = identified.Format;
        }

        if (format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg) {
            pdfBytes = bytes;
            return true;
        }

        return OfficeImagePngConverter.TryConvertToPng(bytes, out pdfBytes);
    }

    private static IEnumerable<string> EnumerateUsedFontFamilyLists(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in EnumerateVisuals(visuals)) {
            if (visual is HtmlRenderText text) {
                yield return text.Font.FamilyName;
            } else if (visual is HtmlRenderDrawing drawing) {
                foreach (string familyNames in EnumerateDrawingFontFamilyLists(drawing.Drawing.Elements)) yield return familyNames;
            }
        }
    }

    private static IEnumerable<string> EnumerateDrawingFontFamilyLists(IEnumerable<OfficeDrawingElement> elements) {
        foreach (OfficeDrawingElement element in elements) {
            if (element is OfficeDrawingText text) {
                yield return text.Font.FamilyName;
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                foreach (string familyNames in EnumerateDrawingFontFamilyLists(effectGroup.Drawing.Elements)) yield return familyNames;
            }
        }
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clipGroup
                ? clipGroup.Visuals
                : visual is HtmlRenderPathClipGroup pathClipGroup
                    ? pathClipGroup.Visuals
                    : visual is HtmlRenderEffectGroup effectGroup ? effectGroup.Visuals
                    : visual is HtmlRenderSemanticGroup semanticGroup ? semanticGroup.Visuals
                    : visual is HtmlRenderLogicalTextGroup logicalTextGroup ? logicalTextGroup.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateVisuals(children)) yield return child;
        }
    }

    internal static PdfCore.PdfTextFallbackFeatures ResolveTextFallbackFeatures(
        HtmlRenderDocument rendered,
        PdfCore.PdfTextFallbackFeatures requested) {
        if (requested == PdfCore.PdfTextFallbackFeatures.None) return requested;

        foreach (HtmlRenderVisual visual in EnumerateVisuals(rendered.Pages.SelectMany(page => page.Visuals))) {
            if (visual is HtmlRenderText text && RequiresUnicodeFont(text.Text)) {
                return requested;
            }

            if (visual is HtmlRenderDrawing drawing && DrawingRequiresUnicodeFont(drawing.Drawing.Elements)) {
                return requested;
            }
        }

        return PdfCore.PdfTextFallbackFeatures.None;
    }

    private static bool DrawingRequiresUnicodeFont(IEnumerable<OfficeDrawingElement> elements) {
        foreach (OfficeDrawingElement element in elements) {
            if (element is OfficeDrawingText text && RequiresUnicodeFont(text.Text)) return true;
            if (element is OfficeDrawingEffectGroup effectGroup && DrawingRequiresUnicodeFont(effectGroup.Drawing.Elements)) return true;
        }

        return false;
    }

    private static bool RequiresUnicodeFont(string text) =>
        PdfCore.PdfTextDiagnostics.AnalyzeWinAnsiText(text).Count != 0;

    private static void RegisterFamily(
        PdfCore.PdfDocument pdf,
        PdfCore.PdfStandardFont slot,
        string family,
        IReadOnlyList<OfficeFontFace> faces,
        CancellationToken cancellationToken) {
        OfficeFontFace regular = FindFace(faces, OfficeFontStyle.Regular) ?? faces[0];
        OfficeFontFace bold = FindFace(faces, OfficeFontStyle.Bold) ?? regular;
        OfficeFontFace italic = FindFace(faces, OfficeFontStyle.Italic) ?? regular;
        OfficeFontFace boldItalic = FindFace(faces, OfficeFontStyle.Bold | OfficeFontStyle.Italic) ?? bold;
        cancellationToken.ThrowIfCancellationRequested();
        EmbedFace(pdf, slot, family, "Regular", regular);
        cancellationToken.ThrowIfCancellationRequested();
        EmbedFace(pdf, PdfCore.PdfStandardFontMapper.GetStyledFont(slot, bold: true, italic: false), family, "Bold", bold);
        cancellationToken.ThrowIfCancellationRequested();
        EmbedFace(pdf, PdfCore.PdfStandardFontMapper.GetStyledFont(slot, bold: false, italic: true), family, "Italic", italic);
        cancellationToken.ThrowIfCancellationRequested();
        EmbedFace(pdf, PdfCore.PdfStandardFontMapper.GetStyledFont(slot, bold: true, italic: true), family, "BoldItalic", boldItalic);
    }

    private static OfficeFontFace? FindFace(IReadOnlyList<OfficeFontFace> faces, OfficeFontStyle style) {
        OfficeFontStyle normalized = style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        return faces.FirstOrDefault(face => (face.Style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic)) == normalized);
    }

    private static void EmbedFace(
        PdfCore.PdfDocument pdf,
        PdfCore.PdfStandardFont slot,
        string family,
        string style,
        OfficeFontFace face) =>
        pdf.EmbedStandardFont(slot, face.Data, family + "-" + style);

    private static IEnumerable<string> EnumerateFamilies(string? familyNames) {
        if (string.IsNullOrWhiteSpace(familyNames)) {
            yield break;
        }

        foreach (string raw in familyNames!.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)) {
            string family = raw.Trim().Trim('"', '\'');
            if (family.Length > 0) {
                yield return family;
            }
        }
    }

    private static PdfCore.PdfAlign MapAlignment(OfficeTextAlignment alignment) {
        if (alignment == OfficeTextAlignment.Center) return PdfCore.PdfAlign.Center;
        if (alignment == OfficeTextAlignment.Right) return PdfCore.PdfAlign.Right;
        if (alignment == OfficeTextAlignment.Justify) return PdfCore.PdfAlign.Justify;
        return PdfCore.PdfAlign.Left;
    }
}
