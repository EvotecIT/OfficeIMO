using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

internal sealed partial class OfficeMarkupPowerPointExporter {
    private static void ExportBlock(
        PowerPointSlide slide,
        OfficeMarkupBlock block,
        MarkupToPowerPointOptions options,
        SlideCanvasMetrics metrics,
        LayoutCursor cursor,
        OfficeMarkupStyleResolver styleResolver) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                AddText(slide, heading.Text, cursor, height: heading.Level <= 2 ? 0.45 : 0.34, styleResolver.Resolve(heading));
                break;
            case OfficeMarkupParagraphBlock paragraph:
                AddText(slide, paragraph.Text, cursor, height: EstimateTextHeight(paragraph.Text), styleResolver.Resolve(paragraph));
                break;
            case OfficeMarkupListBlock list:
                AddList(slide, list, cursor, styleResolver.Resolve("body"));
                break;
            case OfficeMarkupImageBlock image:
                AddImage(slide, image, cursor, options, metrics);
                break;
            case OfficeMarkupTableBlock table:
                AddTable(slide, table, cursor);
                break;
            case OfficeMarkupDiagramBlock diagram:
                AddDiagram(slide, diagram, cursor, options, metrics, styleResolver);
                break;
            case OfficeMarkupChartBlock chart:
                AddChart(slide, chart, cursor, options, metrics);
                break;
            case OfficeMarkupTextBoxBlock textBox:
                AddTextBox(slide, textBox, cursor, metrics, styleResolver);
                break;
            case OfficeMarkupCardBlock card:
                AddCard(slide, card, cursor, metrics, styleResolver);
                break;
            case OfficeMarkupColumnsBlock:
            case OfficeMarkupColumnBlock:
                break;
            case OfficeMarkupExtensionBlock extension:
                ExportExtension(slide, extension, options, metrics, cursor, styleResolver);
                break;
            default:
                if (options.IncludeUnsupportedBlocksAsText) {
                    AddText(slide, block.Kind.ToString(), cursor, height: 0.4, styleResolver.Resolve("caption"));
                }

                break;
        }
    }

    private static void ExportExtension(
        PowerPointSlide slide,
        OfficeMarkupExtensionBlock extension,
        MarkupToPowerPointOptions options,
        SlideCanvasMetrics metrics,
        LayoutCursor cursor,
        OfficeMarkupStyleResolver styleResolver) {
        switch (Normalize(extension.Command)) {
            case "textbox":
                AddTextBox(slide, extension.Body, null, extension.Attributes, cursor, metrics, styleResolver.Resolve(extension));
                break;
            case "card":
                AddCard(slide, GetAttribute(extension.Attributes, "title"), extension.Body, null, extension.Attributes, cursor, metrics, styleResolver.Resolve(extension));
                break;
            case "column":
            case "left":
            case "right":
            case "columns":
                break;
            default:
                if (options.IncludeUnsupportedBlocksAsText && !string.IsNullOrWhiteSpace(extension.Body)) {
                    AddText(slide, extension.Body.Trim(), cursor, height: EstimateTextHeight(extension.Body), styleResolver.Resolve(extension));
                }

                break;
        }
    }

    private static void AddText(PowerPointSlide slide, string text, LayoutCursor cursor, double height, OfficeMarkupResolvedStyle? style = null) {
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        var actualHeight = Math.Max(0.28, Math.Min(height, cursor.RemainingHeight));
        var textBox = slide.AddTextBoxInches(text.Trim(), cursor.Left, cursor.Top, cursor.Width, actualHeight);
        ApplyTextStyle(textBox, style);
        cursor.Advance(actualHeight);
    }

    private static void AddTextBox(PowerPointSlide slide, OfficeMarkupTextBoxBlock textBox, LayoutCursor cursor, SlideCanvasMetrics metrics, OfficeMarkupStyleResolver styleResolver) =>
        AddTextBox(slide, textBox.Text, textBox.Placement, textBox.Attributes, cursor, metrics, styleResolver.Resolve(textBox));

    private static void AddTextBox(
        PowerPointSlide slide,
        string text,
        OfficeMarkupPlacement? placement,
        IDictionary<string, string> attributes,
        LayoutCursor cursor,
        SlideCanvasMetrics metrics,
        OfficeMarkupResolvedStyle? style) {
        var box = ResolveBox(placement, attributes, cursor, 0.62, metrics);
        var textBox = slide.AddTextBoxInches(text.Trim(), box.Left, box.Top, box.Width, box.Height);
        ApplyTextStyle(textBox, style);
        if (!HasExplicitPlacement(placement, attributes)) {
            cursor.Advance(box.Height);
        }
    }

    private static void AddCard(PowerPointSlide slide, OfficeMarkupCardBlock card, LayoutCursor cursor, SlideCanvasMetrics metrics, OfficeMarkupStyleResolver styleResolver) =>
        AddCard(slide, card.Title, card.Body, card.Placement, card.Attributes, cursor, metrics, styleResolver.Resolve(card));

    private static void AddCard(
        PowerPointSlide slide,
        string? title,
        string body,
        OfficeMarkupPlacement? placement,
        IDictionary<string, string> attributes,
        LayoutCursor cursor,
        SlideCanvasMetrics metrics,
        OfficeMarkupResolvedStyle? style) {
        var text = string.IsNullOrWhiteSpace(title)
            ? body.Trim()
            : title!.Trim() + Environment.NewLine + body.Trim();
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        var box = ResolveBox(placement, attributes, cursor, Math.Min(1.25, cursor.RemainingHeight), metrics);
        AddPanel(slide, box, style, "OfficeIMO Markup Card Panel");
        var textBox = slide.AddTextBoxInches(text, box.Left, box.Top, box.Width, box.Height);
        ApplyTextStyle(textBox, style);
        if (!HasExplicitPlacement(placement, attributes)) {
            cursor.Advance(box.Height);
        }
    }

    private static void AddList(PowerPointSlide slide, OfficeMarkupListBlock list, LayoutCursor cursor, OfficeMarkupResolvedStyle? style) {
        var items = list.Items.Select(item => item.Text).Where(text => !string.IsNullOrWhiteSpace(text)).ToList();
        if (items.Count == 0) {
            return;
        }

        var height = Math.Max(0.45, Math.Min(cursor.RemainingHeight, 0.28 * items.Count + 0.25));
        var box = slide.AddTextBoxInches(string.Empty, cursor.Left, cursor.Top, cursor.Width, height);
        box.Clear();
        if (list.Ordered) {
            box.SetNumberedList(items, list.Start);
        } else {
            box.SetBullets(items);
        }

        ApplyTextStyle(box, style);
        cursor.Advance(height);
    }

    private static void AddDiagram(
        PowerPointSlide slide,
        OfficeMarkupDiagramBlock diagram,
        LayoutCursor cursor,
        MarkupToPowerPointOptions options,
        SlideCanvasMetrics metrics,
        OfficeMarkupStyleResolver styleResolver) {
        var box = ResolveBox(diagram.Placement, diagram.Attributes, cursor, Math.Min(2.4, cursor.RemainingHeight), metrics);
        if (ShouldAddVisualPanel(diagram.Attributes, defaultValue: true)) {
            AddVisualPanel(slide, box, metrics, "OfficeIMO Markup Diagram Panel");
        }

        if (OfficeMarkupMermaidRenderer.TryRenderPng(diagram, options, out var imagePath)) {
            try {
                AddPicture(slide, imagePath, box, GetAttribute(diagram.Attributes, "fit"));
                if (!HasExplicitPlacement(diagram.Placement, diagram.Attributes)) {
                    cursor.Advance(box.Height);
                }

                return;
            } finally {
                TryDelete(imagePath);
            }
        }

        if (options.IncludeUnsupportedBlocksAsText) {
            var text = IsMermaid(diagram.Language)
                ? "Mermaid diagram\nInstall or configure the Mermaid renderer to export this block as an image."
                : $"{diagram.Language} diagram";
            var textBox = slide.AddTextBoxInches(text.Trim(), box.Left, box.Top, box.Width, box.Height);
            ApplyTextStyle(textBox, styleResolver.Resolve("caption"));
            if (!HasExplicitPlacement(diagram.Placement, diagram.Attributes)) {
                cursor.Advance(box.Height);
            }
        }
    }
}
