using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class PowerPointHtmlConverterExtensions {
    private static void AppendSemanticShapes(StringBuilder body, PptCore.PowerPointSlide slide, PowerPointHtmlSaveOptions options) {
        foreach (PptCore.PowerPointShape shape in slide.Shapes.OrderBy(shape => shape.DrawingOrder)) {
            if (!options.IncludeHiddenShapes && shape.Hidden) {
                continue;
            }

            if (shape is PptCore.PowerPointTextBox textBox) {
                IReadOnlyList<PptCore.PowerPointParagraph> paragraphs = textBox.Paragraphs;
                if (paragraphs.All(paragraph => paragraph.InlineNodes.All(node => string.IsNullOrEmpty(node.Text)))) {
                    continue;
                }

                body.Append("<p");
                AppendSemanticShapeAttributes(body, textBox, "text");
                body.Append('>');
                AppendSemanticParagraphContent(body, paragraphs, textBox.TextBody?.ListStyle, textBox.MasterTextStyle);
                body.Append("</p>");
            } else if (shape is PptCore.PowerPointTable table && options.IncludeTables) {
                AppendTable(body, table, includeShapeMetadata: true);
            }
        }
    }

    private static void AppendSemanticParagraphContent(
        StringBuilder body,
        IReadOnlyList<PptCore.PowerPointParagraph> paragraphs,
        DocumentFormat.OpenXml.Drawing.ListStyle? listStyle,
        DocumentFormat.OpenXml.OpenXmlCompositeElement? masterTextStyle) {
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
            PptCore.PowerPointParagraph paragraph = paragraphs[paragraphIndex];
            if (paragraphIndex > 0) body.Append("<br data-officeimo-powerpoint-paragraph-break=\"true\">");
            foreach (PptCore.PowerPointParagraphInline node in paragraph.InlineNodes) {
                if (node.Kind == PptCore.PowerPointParagraphInlineKind.Run && node.Run != null) {
                    AppendSemanticTextRun(body, node.Run,
                        PptCore.PowerPointEffectiveRunStyleResolver.Resolve(node.Run, paragraph, listStyle, masterTextStyle));
                } else if (node.Kind == PptCore.PowerPointParagraphInlineKind.LineBreak) {
                    body.Append("<br data-officeimo-powerpoint-inline-break=\"true\">");
                } else if (node.Kind == PptCore.PowerPointParagraphInlineKind.Field) {
                    body.Append("<span data-officeimo-powerpoint-field=\"true\"");
                    if (!string.IsNullOrWhiteSpace(node.FieldId)) {
                        body.Append(" data-officeimo-powerpoint-field-id=\"")
                            .Append(OfficeHtmlText.EscapeAttribute(node.FieldId!))
                            .Append('"');
                    }
                    if (!string.IsNullOrWhiteSpace(node.FieldType)) {
                        body.Append(" data-officeimo-powerpoint-field-type=\"")
                            .Append(OfficeHtmlText.EscapeAttribute(node.FieldType!))
                            .Append('"');
                    }
                    PptCore.PowerPointEffectiveRunStyle effective = default;
                    if (node.Run != null) {
                        effective = PptCore.PowerPointEffectiveRunStyleResolver.Resolve(node.Run, paragraph, listStyle, masterTextStyle);
                        AppendSemanticTextStyleAttributes(body, node.Run, effective);
                    }
                    body.Append('>');
                    if (node.Run != null) AppendSemanticTextRunContent(body, node.Run, effective);
                    else body.Append(OfficeHtmlText.Escape(node.Text));
                    body.Append("</span>");
                }
            }
        }
    }

    private static bool RequiresSemanticTableCellContent(PptCore.PowerPointTableCell cell) {
        IReadOnlyList<PptCore.PowerPointParagraph> paragraphs = cell.Paragraphs;
        if (paragraphs.Count != 1) return true;
        IReadOnlyList<PptCore.PowerPointParagraphInline> nodes = paragraphs[0].InlineNodes;
        if (nodes.Count != 1 || nodes[0].Kind != PptCore.PowerPointParagraphInlineKind.Run || nodes[0].Run == null) {
            return true;
        }

        PptCore.PowerPointTextRun run = nodes[0].Run!;
        DocumentFormat.OpenXml.Drawing.TextBody? textBody = cell.Cell.TextBody;
        DocumentFormat.OpenXml.OpenXmlCompositeElement? masterTextStyle = cell.SlidePart?.SlideLayoutPart?.SlideMasterPart?
            .SlideMaster?.TextStyles?.OtherStyle;
        PptCore.PowerPointEffectiveRunStyle effective = PptCore.PowerPointEffectiveRunStyleResolver.Resolve(
            run, paragraphs[0], textBody?.ListStyle, masterTextStyle);
        return effective.Bold == true || effective.Italic == true
            || effective.UnderlineStyle is { } underline && underline != PptCore.PowerPointUnderlineStyle.None
            || effective.StrikeStyle is { } strike && strike != PptCore.PowerPointStrikeStyle.None
            || effective.Capitalization is { } capitalization && capitalization != PptCore.PowerPointCapitalization.None
            || effective.BaselinePercent.HasValue
            || effective.FontSizePoints.HasValue || !string.IsNullOrWhiteSpace(effective.FontName)
            || !string.IsNullOrWhiteSpace(effective.Color)
            || !string.IsNullOrWhiteSpace(effective.Language)
               && !string.Equals(effective.Language, PptCore.PowerPointTableTextDefaults.Language, StringComparison.OrdinalIgnoreCase)
            || run.Hyperlink != null;
    }

    private static void AppendSemanticTextRun(
        StringBuilder body,
        PptCore.PowerPointTextRun run,
        PptCore.PowerPointEffectiveRunStyle effective) {
        body.Append("<span data-officeimo-powerpoint-run=\"true\"");
        AppendSemanticTextStyleAttributes(body, run, effective);
        body.Append('>');
        AppendSemanticTextRunContent(body, run, effective);
        body.Append("</span>");
    }

    private static void AppendSemanticTextStyleAttributes(
        StringBuilder body,
        PptCore.PowerPointTextRun run,
        PptCore.PowerPointEffectiveRunStyle effective) {
        var css = new StringBuilder();
        AppendCss(css, "font-weight", effective.Bold == true ? "700" : null);
        AppendCss(css, "font-style", effective.Italic == true ? "italic" : null);
        AppendCss(css, "font-family", !string.IsNullOrWhiteSpace(effective.FontName)
            ? "'" + effective.FontName!.Replace("'", "\\'") + "'" : null);
        AppendCss(css, "font-size", effective.FontSizePoints.HasValue
            ? effective.FontSizePoints.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt" : null);
        AppendCss(css, "color", !string.IsNullOrWhiteSpace(effective.Color) ? "#" + effective.Color!.TrimStart('#') : null);
        AppendCss(css, "vertical-align", effective.BaselinePercent switch {
            > 0D => "super",
            < 0D => "sub",
            _ => null
        });
        AppendCss(css, "font-variant", effective.Capitalization == PptCore.PowerPointCapitalization.SmallCaps ? "small-caps" : null);
        AppendCss(css, "text-transform", effective.Capitalization == PptCore.PowerPointCapitalization.AllCaps ? "uppercase" : null);

        if (css.Length > 0) {
            body.Append(" style=\"")
                .Append(OfficeHtmlText.EscapeAttribute(css.ToString()))
                .Append('"');
        }
        if (effective.UnderlineStyle.HasValue) {
            body.Append(" data-officeimo-powerpoint-underline=\"")
                .Append(OfficeHtmlText.EscapeAttribute(effective.UnderlineStyle.Value.ToString()))
                .Append('"');
        }
        if (effective.StrikeStyle.HasValue) {
            body.Append(" data-officeimo-powerpoint-strike=\"")
                .Append(OfficeHtmlText.EscapeAttribute(effective.StrikeStyle.Value.ToString()))
                .Append('"');
        }
        if (effective.Capitalization.HasValue) {
            body.Append(" data-officeimo-powerpoint-capitalization=\"")
                .Append(OfficeHtmlText.EscapeAttribute(effective.Capitalization.Value.ToString()))
                .Append('"');
        }
        if (effective.BaselinePercent.HasValue) {
            body.Append(" data-officeimo-powerpoint-baseline-percent=\"")
                .Append(effective.BaselinePercent.Value.ToString("0.###", CultureInfo.InvariantCulture))
                .Append('"');
        }
        if (!string.IsNullOrWhiteSpace(effective.FontName)) {
            body.Append(" data-officeimo-powerpoint-font-family=\"")
                .Append(OfficeHtmlText.EscapeAttribute(effective.FontName!))
                .Append('"');
        }
        if (!string.IsNullOrWhiteSpace(effective.Language)) {
            body.Append(" data-officeimo-powerpoint-language=\"")
                .Append(OfficeHtmlText.EscapeAttribute(effective.Language!))
                .Append('"');
        }
        if (run.Hyperlink != null) {
            body.Append(" data-officeimo-powerpoint-hyperlink=\"")
                .Append(OfficeHtmlText.EscapeAttribute(run.Hyperlink.ToString()))
                .Append('"');
        }
    }

    private static void AppendSemanticTextRunContent(
        StringBuilder body,
        PptCore.PowerPointTextRun run,
        PptCore.PowerPointEffectiveRunStyle effective) {
        AppendPowerPointDecorationStart(body, "underline", GetPowerPointUnderlineCssStyle(effective.UnderlineStyle));
        AppendPowerPointDecorationStart(body, "line-through", GetPowerPointStrikeCssStyle(effective.StrikeStyle));
        body.Append(OfficeHtmlText.Escape(run.Text));
        AppendPowerPointDecorationEnd(body, effective.StrikeStyle.HasValue && effective.StrikeStyle.Value != PptCore.PowerPointStrikeStyle.None);
        AppendPowerPointDecorationEnd(body, effective.UnderlineStyle.HasValue && effective.UnderlineStyle.Value != PptCore.PowerPointUnderlineStyle.None);
    }

    private static string? GetPowerPointUnderlineCssStyle(PptCore.PowerPointUnderlineStyle? underline) => underline switch {
            PptCore.PowerPointUnderlineStyle.Double or PptCore.PowerPointUnderlineStyle.WavyDouble => "double",
            PptCore.PowerPointUnderlineStyle.Dotted or PptCore.PowerPointUnderlineStyle.HeavyDotted => "dotted",
            PptCore.PowerPointUnderlineStyle.Dash or PptCore.PowerPointUnderlineStyle.DashHeavy or
                PptCore.PowerPointUnderlineStyle.DashLong or PptCore.PowerPointUnderlineStyle.DashLongHeavy or
                PptCore.PowerPointUnderlineStyle.DotDash or PptCore.PowerPointUnderlineStyle.DotDashHeavy or
                PptCore.PowerPointUnderlineStyle.DotDotDash or PptCore.PowerPointUnderlineStyle.DotDotDashHeavy => "dashed",
            PptCore.PowerPointUnderlineStyle.Wavy or PptCore.PowerPointUnderlineStyle.WavyHeavy => "wavy",
            PptCore.PowerPointUnderlineStyle.None or null => null,
            _ => "solid"
        };

    private static string? GetPowerPointStrikeCssStyle(PptCore.PowerPointStrikeStyle? strike) => strike switch {
        PptCore.PowerPointStrikeStyle.Double => "double",
        PptCore.PowerPointStrikeStyle.Single => "solid",
        _ => null
    };

    private static void AppendPowerPointDecorationStart(StringBuilder body, string line, string? style) {
        if (style == null) return;
        body.Append("<span style=\"text-decoration-line:")
            .Append(line)
            .Append(";text-decoration-style:")
            .Append(style)
            .Append("\">");
    }

    private static void AppendPowerPointDecorationEnd(StringBuilder body, bool enabled) {
        if (enabled) body.Append("</span>");
    }

    private static void AppendCss(StringBuilder css, string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) return;
        if (css.Length > 0) css.Append(';');
        css.Append(name).Append(':').Append(value);
    }

    private static void AppendSemanticShapeAttributes(StringBuilder body, PptCore.PowerPointShape shape, string kind) {
        body.Append(" data-officeimo-layer-kind=\"")
            .Append(kind)
            .Append("\" data-officeimo-layer-index=\"")
            .Append(shape.DrawingOrder.ToString(CultureInfo.InvariantCulture))
            .Append('"');
        AppendDataAttribute(body, "data-officeimo-left", shape.LeftPoints, omitWhenZero: false);
        AppendDataAttribute(body, "data-officeimo-top", shape.TopPoints, omitWhenZero: false);
        AppendDataAttribute(body, "data-officeimo-width", shape.WidthPoints, omitWhenZero: false);
        AppendDataAttribute(body, "data-officeimo-height", shape.HeightPoints, omitWhenZero: false);
        AppendDataAttribute(body, "data-officeimo-rotation", shape.Rotation ?? 0D);
        AppendDataAttribute(body, "data-officeimo-flip-horizontal", shape.HorizontalFlip == true);
        AppendDataAttribute(body, "data-officeimo-flip-vertical", shape.VerticalFlip == true);
        AppendDataAttribute(body, "data-officeimo-hidden", shape.Hidden);
    }
}
