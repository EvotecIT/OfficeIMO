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
                for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
                    if (paragraphIndex > 0) body.Append("<br data-officeimo-powerpoint-paragraph-break=\"true\">");
                    PptCore.PowerPointParagraph paragraph = paragraphs[paragraphIndex];
                    foreach (PptCore.PowerPointParagraphInline node in paragraph.InlineNodes) {
                        if (node.Kind == PptCore.PowerPointParagraphInlineKind.Run && node.Run != null) {
                            AppendSemanticTextRun(body, node.Run);
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
                            body.Append('>')
                                .Append(OfficeHtmlText.Escape(node.Text))
                                .Append("</span>");
                        }
                    }
                }
                body.Append("</p>");
            } else if (shape is PptCore.PowerPointTable table && options.IncludeTables) {
                AppendTable(body, table, includeShapeMetadata: true);
            }
        }
    }

    private static void AppendSemanticTextRun(StringBuilder body, PptCore.PowerPointTextRun run) {
        var css = new StringBuilder();
        AppendCss(css, "font-weight", run.Bold ? "700" : null);
        AppendCss(css, "font-style", run.Italic ? "italic" : null);
        AppendCss(css, "font-family", !string.IsNullOrWhiteSpace(run.FontName)
            ? "'" + run.FontName!.Replace("'", "\\'") + "'" : null);
        AppendCss(css, "font-size", run.FontSizePoints.HasValue
            ? run.FontSizePoints.Value.ToString("0.###", CultureInfo.InvariantCulture) + "pt" : null);
        AppendCss(css, "color", !string.IsNullOrWhiteSpace(run.Color) ? "#" + run.Color!.TrimStart('#') : null);
        AppendCss(css, "vertical-align", run.BaselinePercent switch {
            > 0D => "super",
            < 0D => "sub",
            _ => null
        });
        AppendCss(css, "font-variant", run.Capitalization == PptCore.PowerPointCapitalization.SmallCaps ? "small-caps" : null);
        AppendCss(css, "text-transform", run.Capitalization == PptCore.PowerPointCapitalization.AllCaps ? "uppercase" : null);

        body.Append("<span");
        if (css.Length > 0) {
            body.Append(" style=\"")
                .Append(OfficeHtmlText.EscapeAttribute(css.ToString()))
                .Append('"');
        }
        if (run.UnderlineStyle.HasValue) {
            body.Append(" data-officeimo-powerpoint-underline=\"")
                .Append(OfficeHtmlText.EscapeAttribute(run.UnderlineStyle.Value.ToString()))
                .Append('"');
        }
        if (run.StrikeStyle.HasValue) {
            body.Append(" data-officeimo-powerpoint-strike=\"")
                .Append(OfficeHtmlText.EscapeAttribute(run.StrikeStyle.Value.ToString()))
                .Append('"');
        }
        if (run.Capitalization.HasValue) {
            body.Append(" data-officeimo-powerpoint-capitalization=\"")
                .Append(OfficeHtmlText.EscapeAttribute(run.Capitalization.Value.ToString()))
                .Append('"');
        }
        if (run.BaselinePercent.HasValue) {
            body.Append(" data-officeimo-powerpoint-baseline-percent=\"")
                .Append(run.BaselinePercent.Value.ToString("0.###", CultureInfo.InvariantCulture))
                .Append('"');
        }
        if (!string.IsNullOrWhiteSpace(run.FontName)) {
            body.Append(" data-officeimo-powerpoint-font-family=\"")
                .Append(OfficeHtmlText.EscapeAttribute(run.FontName!))
                .Append('"');
        }
        body.Append('>');
        AppendPowerPointDecorationStart(body, "underline", GetPowerPointUnderlineCssStyle(run.UnderlineStyle));
        AppendPowerPointDecorationStart(body, "line-through", GetPowerPointStrikeCssStyle(run.StrikeStyle));
        body.Append(OfficeHtmlText.Escape(run.Text));
        AppendPowerPointDecorationEnd(body, run.StrikeStyle.HasValue && run.StrikeStyle.Value != PptCore.PowerPointStrikeStyle.None);
        AppendPowerPointDecorationEnd(body, run.UnderlineStyle.HasValue && run.UnderlineStyle.Value != PptCore.PowerPointUnderlineStyle.None);
        body.Append("</span>");
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
