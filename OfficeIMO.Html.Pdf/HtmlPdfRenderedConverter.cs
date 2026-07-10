using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static class HtmlPdfRenderedConverter {
    private const double PointsPerCssPixel = 72D / HtmlRenderOptions.CssPixelsPerInch;

    internal static PdfCore.PdfDocument Convert(string html, HtmlPdfSaveOptions options) {
        HtmlRenderOptions renderOptions = options.RenderOptions?.Clone() ?? new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged
        };
        options.RenderOptions = renderOptions;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, renderOptions);
        return CreatePdf(rendered, options);
    }

    internal static async Task<PdfCore.PdfDocument> ConvertAsync(string html, HtmlPdfSaveOptions options, CancellationToken cancellationToken) {
        HtmlRenderOptions renderOptions = options.RenderOptions?.Clone() ?? new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged
        };
        options.RenderOptions = renderOptions;
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(html, renderOptions, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return CreatePdf(rendered, options);
    }

    private static PdfCore.PdfDocument CreatePdf(HtmlRenderDocument rendered, HtmlPdfSaveOptions options) {
        options.RenderDiagnostics = rendered.Diagnostics.Clone();

        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create();
        if (options.RenderedFontFamily != null) {
            pdf.UseFontFamily(options.RenderedFontFamily);
        }

        pdf.UseTextFallbacks(options.RenderedTextFallbacks)
            .UseTextShaping(options.RenderedTextShapingMode, options.RenderedTextShapingProvider);
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts = RegisterWebFonts(pdf, rendered, options.RenderDiagnostics);
        foreach (HtmlRenderPage renderedPage in rendered.Pages) {
            double pageWidth = renderedPage.Width * PointsPerCssPixel;
            double pageHeight = renderedPage.Height * PointsPerCssPixel;
            pdf.Page(page => page
                .Size(pageWidth, pageHeight)
                .Margin(0D)
                .Canvas(canvas => AddPageVisuals(canvas, renderedPage, webFonts)));
        }

        return pdf;
    }

    private static void AddPageVisuals(PdfCore.PdfPageCanvas canvas, HtmlRenderPage page, IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts) {
        foreach (HtmlRenderVisual visual in page.Visuals.OrderBy(item => item.PaintOrder)) {
            AddVisual(canvas, visual, webFonts, page.Width, page.Height);
        }
    }

    private static void AddVisual(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderVisual visual,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        double surfaceWidth,
        double surfaceHeight) {
        if (visual is HtmlRenderShape shape) {
            AddShape(canvas, shape);
        } else if (visual is HtmlRenderText text) {
            AddText(canvas, text, webFonts);
        } else if (visual is HtmlRenderImage image) {
            AddImage(canvas, image);
        } else if (visual is HtmlRenderImagePattern imagePattern) {
            AddImagePattern(canvas, imagePattern);
        } else if (visual is HtmlRenderClipGroup group) {
            AddClipGroup(canvas, group, webFonts, surfaceWidth, surfaceHeight);
        } else if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            AddPathClipGroup(canvas, pathClipGroup, webFonts, surfaceWidth, surfaceHeight);
        } else if (visual is HtmlRenderEffectGroup effectGroup) {
            AddEffectGroup(canvas, effectGroup, webFonts, surfaceWidth, surfaceHeight);
        }
    }

    private static void AddEffectGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderEffectGroup group,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        double surfaceWidth,
        double surfaceHeight) {
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
                AddVisual(nested, child, webFonts, surfaceWidth, surfaceHeight);
            }
        });
    }

    private static void AddClipGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderClipGroup group,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        double surfaceWidth,
        double surfaceHeight) {
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
                    AddVisual(clipped, child, webFonts, surfaceWidth, surfaceHeight);
                }
            });
    }

    private static void AddPathClipGroup(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderPathClipGroup group,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts,
        double surfaceWidth,
        double surfaceHeight) {
        canvas.Clip(
            group.ClipX * PointsPerCssPixel,
            group.ClipY * PointsPerCssPixel,
            group.ClipPath.Scale(PointsPerCssPixel, PointsPerCssPixel),
            clipped => {
                foreach (HtmlRenderVisual child in group.Visuals.OrderBy(item => item.PaintOrder)) {
                    AddVisual(clipped, child, webFonts, surfaceWidth, surfaceHeight);
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

    private static void AddText(PdfCore.PdfPageCanvas canvas, HtmlRenderText visual, IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts) {
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
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            MapAlignment(visual.Alignment),
            visual.Font.Size * PointsPerCssPixel,
            visual.LineHeight * PointsPerCssPixel);
    }

    private static void AddImage(PdfCore.PdfPageCanvas canvas, HtmlRenderImage visual) {
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
            visual.Bytes,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            style,
            linkUri: visual.LinkUri,
            linkContents: visual.LinkUri == null ? null : visual.Source,
            alternativeText: visual.AlternativeText);
    }

    private static void AddImagePattern(PdfCore.PdfPageCanvas canvas, HtmlRenderImagePattern visual) {
        OfficeImagePatternLayout pattern = visual.Pattern.Scale(PointsPerCssPixel);
        OfficeImagePlacement area = pattern.Area;
        PdfCore.PdfCanvasImageResource imageResource = PdfCore.PdfCanvasImageResource.Create(visual.Bytes);
        canvas.Clip(area.X, area.Y, area.Width, area.Height, clipped => {
            foreach (OfficeImagePlacement tile in pattern.GetTilePlacements(visual.MaximumTileCount)) {
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

        string normalized = familyName ?? string.Empty;
        if (normalized.IndexOf("times", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("serif", StringComparison.OrdinalIgnoreCase) >= 0) {
            return PdfCore.PdfStandardFont.TimesRoman;
        }

        if (normalized.IndexOf("courier", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("consolas", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("mono", StringComparison.OrdinalIgnoreCase) >= 0) {
            return PdfCore.PdfStandardFont.Courier;
        }

        return PdfCore.PdfStandardFont.Helvetica;
    }

    private static IReadOnlyDictionary<string, PdfCore.PdfStandardFont> RegisterWebFonts(
        PdfCore.PdfDocument pdf,
        HtmlRenderDocument rendered,
        HtmlDiagnosticReport? diagnostics) {
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
        foreach (HtmlRenderText text in rendered.Pages.SelectMany(page => EnumerateVisuals(page.Visuals)).OfType<HtmlRenderText>()) {
            foreach (string family in EnumerateFamilies(text.Font.FamilyName)) {
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
        for (int index = 0; index < orderedFamilies.Count; index++) {
            string family = orderedFamilies[index];
            if (index >= slots.Length) {
                diagnostics?.Add(
                    "OfficeIMO.Html.Pdf",
                    HtmlPdfDiagnosticCodes.RenderedFontFamilyLimitExceeded,
                    "The rendered PDF can embed three distinct active web-font families; an additional family used standard-font fallback.",
                    HtmlDiagnosticSeverity.Warning,
                    family,
                    "limit=" + slots.Length);
                continue;
            }

            PdfCore.PdfStandardFont slot = slots[index];
            RegisterFamily(pdf, slot, family, byFamily[family]);
            mappings[family] = slot;
        }

        return mappings;
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clipGroup
                ? clipGroup.Visuals
                : visual is HtmlRenderEffectGroup effectGroup ? effectGroup.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateVisuals(children)) yield return child;
        }
    }

    private static void RegisterFamily(
        PdfCore.PdfDocument pdf,
        PdfCore.PdfStandardFont slot,
        string family,
        IReadOnlyList<OfficeFontFace> faces) {
        OfficeFontFace regular = FindFace(faces, OfficeFontStyle.Regular) ?? faces[0];
        OfficeFontFace bold = FindFace(faces, OfficeFontStyle.Bold) ?? regular;
        OfficeFontFace italic = FindFace(faces, OfficeFontStyle.Italic) ?? regular;
        OfficeFontFace boldItalic = FindFace(faces, OfficeFontStyle.Bold | OfficeFontStyle.Italic) ?? bold;
        EmbedFace(pdf, slot, family, "Regular", regular);
        EmbedFace(pdf, PdfCore.PdfStandardFontMapper.GetStyledFont(slot, bold: true, italic: false), family, "Bold", bold);
        EmbedFace(pdf, PdfCore.PdfStandardFontMapper.GetStyledFont(slot, bold: false, italic: true), family, "Italic", italic);
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
