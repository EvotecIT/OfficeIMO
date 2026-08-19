using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class PowerPointHtmlConverterExtensions {
    private static void AppendMasterInventory(StringBuilder body, PptCore.PowerPointPresentation presentation,
        IList<HtmlDiagnostic> diagnostics) {
        try {
            PptCore.PowerPointTemplateInventory inventory = PptCore.PowerPointTemplate.Inspect(presentation);
            if (inventory.Masters.Count == 0) return;
            body.Append("<section class=\"officeimo-feature officeimo-masters\"><h2>Presentation masters</h2>")
                .Append("<div class=\"officeimo-diagnostic\" data-officeimo-loss=\"review-only\">")
                .Append("Masters and layouts are exposed as inert template metadata; HTML does not execute PowerPoint inheritance or theme editing behavior.")
                .Append("</div><ul class=\"officeimo-feature-list\">");
            foreach (PptCore.PowerPointTemplateMasterInfo master in inventory.Masters) {
                diagnostics.Add(new HtmlDiagnostic(
                    "OfficeIMO.PowerPoint.Html",
                    HtmlConversionDiagnosticCodes.PowerPointMasterReviewApproximated,
                    "Slide master '" + (string.IsNullOrWhiteSpace(master.Name) ? (master.MasterIndex + 1).ToString(CultureInfo.InvariantCulture) : master.Name) + "' and its layouts were exported as inert template metadata; PowerPoint inheritance and theme editing remain native behavior.",
                    HtmlDiagnosticSeverity.Warning,
                    "powerpoint:master:" + master.MasterIndex.ToString(CultureInfo.InvariantCulture),
                    lossKind: OfficeConversionLossKind.Approximation));
                body.Append("<li class=\"officeimo-feature-item\" data-officeimo-feature=\"slide-master\" data-officeimo-master-index=\"")
                    .Append(master.MasterIndex.ToString(CultureInfo.InvariantCulture)).Append("\">")
                    .Append("<span class=\"officeimo-feature-label\">")
                    .Append(OfficeHtmlText.Escape(string.IsNullOrWhiteSpace(master.Name) ? "Master " + (master.MasterIndex + 1) : master.Name))
                    .Append("</span><div class=\"officeimo-feature-meta\">Theme: ")
                    .Append(OfficeHtmlText.Escape(master.ThemeName))
                    .Append("; Layouts: ")
                    .Append(master.Layouts.Count.ToString(CultureInfo.InvariantCulture))
                    .Append("</div><ul>");
                foreach (PptCore.PowerPointTemplateLayoutInfo layout in master.Layouts) {
                    body.Append("<li data-officeimo-feature=\"slide-layout\" data-officeimo-layout-index=\"")
                        .Append(layout.LayoutIndex.ToString(CultureInfo.InvariantCulture)).Append("\">")
                        .Append(OfficeHtmlText.Escape(string.IsNullOrWhiteSpace(layout.Name) ? layout.Type?.ToString() ?? "Layout" : layout.Name))
                        .Append(" <span class=\"officeimo-feature-meta\">(")
                        .Append(layout.Placeholders.Count.ToString(CultureInfo.InvariantCulture))
                        .Append(" placeholders)</span></li>");
                }
                body.Append("</ul></li>");
            }
            body.Append("</ul></section>");
        } catch (Exception ex) when (ex is InvalidDataException || ex is InvalidOperationException || ex is ArgumentException) {
            diagnostics.Add(new HtmlDiagnostic(
                "OfficeIMO.PowerPoint.Html",
                HtmlConversionDiagnosticCodes.PowerPointMasterReviewOmitted,
                "PowerPoint master inventory could not be read safely and was omitted.",
                HtmlDiagnosticSeverity.Warning,
                "powerpoint:masters",
                ex.Message,
                OfficeConversionLossKind.Omission));
            body.Append("<div class=\"officeimo-diagnostic\" data-officeimo-feature=\"slide-master\" data-officeimo-loss=\"omitted\">")
                .Append("Master inventory could not be read safely: ")
                .Append(OfficeHtmlText.Escape(ex.Message)).Append("</div>");
        }
    }

    private static void AppendAdvancedReviewInventory(StringBuilder body, PptCore.PowerPointSlide slide,
        PowerPointHtmlSaveOptions options, IList<HtmlDiagnostic> diagnostics) {
        if (options.IncludeSmartArt) {
            IEnumerable<PptCore.PowerPointSmartArt> smartArts = options.IncludeHiddenShapes
                ? slide.SmartArts
                : slide.SmartArts.Where(item => !item.Hidden);
            AppendSmartArtInventory(body, smartArts, diagnostics);
        }
        if (options.IncludeMedia) {
            IEnumerable<PptCore.PowerPointMedia> media = options.IncludeHiddenShapes
                ? slide.Media
                : slide.Media.Where(item => !item.Hidden);
            AppendMediaInventory(body, media, options.IncludeAdvancedEffects, diagnostics);
        }
    }

    private static void AppendSmartArtInventory(StringBuilder body, IEnumerable<PptCore.PowerPointSmartArt> smartArts,
        IList<HtmlDiagnostic> diagnostics) {
        List<PptCore.PowerPointSmartArt> items = smartArts.ToList();
        if (items.Count == 0) return;
        body.Append("<section class=\"officeimo-feature officeimo-smartart\"><h3>SmartArt</h3><ul class=\"officeimo-feature-list\">");
        foreach (PptCore.PowerPointSmartArt smartArt in items) {
            diagnostics.Add(new HtmlDiagnostic(
                "OfficeIMO.PowerPoint.Html",
                HtmlConversionDiagnosticCodes.PowerPointSmartArtReviewApproximated,
                "SmartArt '" + GetShapeLabel(smartArt) + "' was exported through a static semantic drawing or text fallback; editable diagram layout remains native PowerPoint behavior.",
                HtmlDiagnosticSeverity.Warning,
                "powerpoint:smartart:" + smartArt.DrawingOrder.ToString(CultureInfo.InvariantCulture),
                lossKind: OfficeConversionLossKind.Approximation));
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-feature=\"smartart\" data-officeimo-layer-index=\"")
                .Append(smartArt.DrawingOrder.ToString(CultureInfo.InvariantCulture)).Append("\">");
            AppendSmartArt(body, smartArt);
            body.Append("</li>");
        }
        body.Append("</ul></section>");
    }

    private static void AppendSmartArt(StringBuilder body, PptCore.PowerPointSmartArt smartArt) {
        string label = GetShapeLabel(smartArt);
        body.Append("<span class=\"officeimo-feature-label\">").Append(OfficeHtmlText.Escape(label)).Append("</span>");
        if (!smartArt.TryGetOfficeDiagramSnapshot(out OfficeDiagramSnapshot snapshot)) {
            IReadOnlyList<string> nodeTexts;
            try { nodeTexts = smartArt.GetNodeTexts(); } catch { nodeTexts = Array.Empty<string>(); }
            body.Append("<div class=\"officeimo-diagnostic\" data-officeimo-loss=\"simplified\">")
                .Append("SmartArt layout is preservation-only; editable node text is shown without invented geometry.</div>");
            AppendTextItems(body, nodeTexts);
            return;
        }

        body.Append("<div class=\"officeimo-feature-meta\">Semantic family: ")
            .Append(OfficeHtmlText.Escape(snapshot.Kind.ToString())).Append("</div>");
        try {
            OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(snapshot);
            body.Append("<div class=\"officeimo-smartart-rendered\" data-officeimo-visual-owner=\"OfficeIMO.Core\">")
                .Append(OfficeDrawingSvgExporter.ToSvg(drawing, 1D, OfficeSvgSizeUnit.Point, null,
                    "officeimo-smartart-" + smartArt.DrawingOrder.ToString(CultureInfo.InvariantCulture) + "-"))
                .Append("</div>");
        } catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException) {
            body.Append("<div class=\"officeimo-diagnostic\" data-officeimo-loss=\"simplified\">")
                .Append("SmartArt semantic snapshot could not be rendered safely: ")
                .Append(OfficeHtmlText.Escape(ex.Message)).Append("</div>");
            AppendTextItems(body, snapshot.Nodes);
        }
    }

    private static void AppendTextItems(StringBuilder body, IEnumerable<string> values) {
        List<string> items = values.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
        if (items.Count == 0) return;
        body.Append("<ul>");
        foreach (string value in items) body.Append("<li>").Append(OfficeHtmlText.Escape(value)).Append("</li>");
        body.Append("</ul>");
    }

    private static void AppendMediaInventory(StringBuilder body, IEnumerable<PptCore.PowerPointMedia> media,
        bool includeAdvancedEffects, IList<HtmlDiagnostic> diagnostics) {
        List<PptCore.PowerPointMedia> items = media.ToList();
        if (items.Count == 0) return;
        body.Append("<section class=\"officeimo-feature officeimo-media\"><h3>Media</h3>")
            .Append("<div class=\"officeimo-diagnostic\" data-officeimo-loss=\"review-only\">")
            .Append("Audio and video are never executed; poster frames and inert playback metadata are emitted for review.</div>")
            .Append("<ul class=\"officeimo-feature-list\">");
        foreach (PptCore.PowerPointMedia item in items) {
            diagnostics.Add(new HtmlDiagnostic(
                "OfficeIMO.PowerPoint.Html",
                HtmlConversionDiagnosticCodes.PowerPointMediaReviewApproximated,
                item.Kind + " media was exported as an inert poster frame and playback inventory; HTML does not execute the media.",
                HtmlDiagnosticSeverity.Warning,
                "powerpoint:media:" + item.DrawingOrder.ToString(CultureInfo.InvariantCulture),
                lossKind: OfficeConversionLossKind.Approximation));
            if (includeAdvancedEffects) AddPictureEffectDiagnostic(item, diagnostics);
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-feature=\"media\" data-officeimo-media-kind=\"")
                .Append(OfficeHtmlText.EscapeAttribute(item.Kind.ToString()))
                .Append("\" data-officeimo-media-source=\"")
                .Append(OfficeHtmlText.EscapeAttribute(item.SourceKind.ToString())).Append('"');
            if (includeAdvancedEffects) AppendPictureEffectAttributes(body, item);
            body.Append('>');
            AppendMediaPoster(body, item);
            body.Append("</li>");
        }
        body.Append("</ul></section>");
    }

    private static void AppendMediaPoster(StringBuilder body, PptCore.PowerPointMedia media) {
        PptCore.PowerPointMediaPlaybackOptions playback;
        try { playback = media.GetPlaybackOptions(); }
        catch { playback = new PptCore.PowerPointMediaPlaybackOptions(); }
        body.Append("<figure data-officeimo-poster-frame=\"true\">");
        AppendPictureShape(body, media);
        body.Append("<figcaption>").Append(OfficeHtmlText.Escape(media.Kind + " poster frame"))
            .Append("; source: ").Append(OfficeHtmlText.Escape(media.SourceKind.ToString()))
            .Append("; volume: ").Append(playback.VolumePercent.ToString(CultureInfo.InvariantCulture)).Append('%');
        if (playback.Mute) body.Append("; muted");
        if (playback.Loop) body.Append("; loop");
        if (playback.FullScreen) body.Append("; full screen");
        body.Append("</figcaption></figure>");
    }

    private static void AppendPictureEffectAttributes(StringBuilder body, PptCore.PowerPointPicture picture) {
        AppendOptionalAttribute(body, "data-officeimo-brightness", picture.LuminanceBrightness);
        AppendOptionalAttribute(body, "data-officeimo-contrast", picture.LuminanceContrast);
        AppendDataAttribute(body, "data-officeimo-grayscale", picture.GrayScale);
        AppendOptionalAttribute(body, "data-officeimo-bilevel", picture.BlackWhiteThreshold);
        if (picture.TransparentColor.HasValue) {
            body.Append(" data-officeimo-transparent-color=\"")
                .Append(OfficeHtmlText.EscapeAttribute(picture.TransparentColor.Value.ToString())).Append('"');
        }
        if (picture.RecolorColor.HasValue) {
            body.Append(" data-officeimo-recolor=\"")
                .Append(OfficeHtmlText.EscapeAttribute(picture.RecolorColor.Value.ToString())).Append('"');
        }
    }

    private static void AppendOptionalAttribute(StringBuilder body, string name, int? value) {
        if (!value.HasValue) return;
        body.Append(' ').Append(name).Append("=\"")
            .Append(value.Value.ToString(CultureInfo.InvariantCulture)).Append('"');
    }

    private static void AddPictureEffectDiagnostic(PptCore.PowerPointPicture picture,
        IList<HtmlDiagnostic> diagnostics) {
        if (!HasReviewEffect(picture)) return;
        diagnostics.Add(new HtmlDiagnostic(
            "OfficeIMO.PowerPoint.Html",
            HtmlConversionDiagnosticCodes.PowerPointEffectReviewApproximated,
            "Picture effects for '" + GetShapeLabel(picture) + "' were preserved as inert review metadata rather than editable PowerPoint effects.",
            HtmlDiagnosticSeverity.Warning,
            "powerpoint:effect:" + picture.DrawingOrder.ToString(CultureInfo.InvariantCulture),
            lossKind: OfficeConversionLossKind.Approximation));
    }

    private static bool HasReviewEffect(PptCore.PowerPointPicture picture) =>
        picture.LuminanceBrightness.HasValue ||
        picture.LuminanceContrast.HasValue ||
        picture.GrayScale ||
        picture.BlackWhiteThreshold.HasValue ||
        picture.TransparentColor.HasValue ||
        picture.RecolorColor.HasValue;
}
