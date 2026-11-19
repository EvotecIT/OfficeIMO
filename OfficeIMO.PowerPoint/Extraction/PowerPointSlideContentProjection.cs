using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// Describes the text boxes and tables of a loaded slide with the content records <see cref="PowerPointSlideReader"/>
/// returns, so extraction renders a loaded presentation and a streamed one through the same code. The records are read
/// straight from the slide's elements and carry only what extraction reads: other shapes are skipped, and identity,
/// visibility and image fields are left empty because the shape enumeration has already applied visibility.
/// </summary>
internal static class PowerPointSlideContentProjection {
    internal static IEnumerable<PowerPointShapeContent> ProjectShapes(PowerPointSlide slide, bool includeHidden) {
        foreach (PowerPointShape shape in slide.EnumerateShapesDeep(slide.Shapes, includeHidden)) {
            if (shape is PowerPointTextBox textBox) {
                yield return new PowerPointShapeContent(PowerPointShapeContentType.TextBox, null, null, null, false,
                    isVisible: true, depth: 0, textBox.ShapePlaceholderType, ProjectParagraphs(textBox), null, null);
            } else if (shape is PowerPointTable table) {
                yield return new PowerPointShapeContent(PowerPointShapeContentType.Table, null, null, null, false,
                    isVisible: true, depth: 0, null, Array.Empty<PowerPointParagraphContent>(), ProjectTable(table, slide.SlidePart), null);
            }
        }
    }

    /// <summary>Projects a loaded text box's paragraphs, as <see cref="PowerPointTextBox.Paragraphs"/> lists them.</summary>
    internal static IReadOnlyList<PowerPointParagraphContent> ProjectParagraphs(PowerPointTextBox textBox) =>
        ProjectParagraphs(textBox.TextBody, textBox.OwnerSlide?.SlidePart);

    private static PowerPointTableContent ProjectTable(PowerPointTable table, SlidePart slidePart) {
        A.Table element = table.TableElement;
        A.TableGrid? grid = element.TableGrid;
        var rows = new List<PowerPointTableRowContent>();
        foreach (A.TableRow row in element.Elements<A.TableRow>()) {
            var cells = new List<PowerPointTableCellContent>();
            foreach (A.TableCell cell in row.Elements<A.TableCell>()) {
                cells.Add(new PowerPointTableCellContent(ProjectParagraphs(cell.TextBody, slidePart)));
            }
            rows.Add(new PowerPointTableRowContent(cells));
        }
        return new PowerPointTableContent(null, grid?.Elements<A.GridColumn>().Count() ?? 0,
            element.TableProperties?.FirstRow?.Value == true, rows, hasGrid: grid != null);
    }

    private static IReadOnlyList<PowerPointParagraphContent> ProjectParagraphs(OpenXmlElement? body, SlidePart? slidePart) {
        if (body == null) return Array.Empty<PowerPointParagraphContent>();
        List<PowerPointParagraphContent>? paragraphs = null;
        foreach (A.Paragraph paragraph in body.Elements<A.Paragraph>()) {
            (paragraphs ??= new List<PowerPointParagraphContent>()).Add(ProjectParagraph(paragraph, slidePart));
        }
        return paragraphs ?? (IReadOnlyList<PowerPointParagraphContent>)Array.Empty<PowerPointParagraphContent>();
    }

    // Reads what PowerPointParagraph and PowerPointTextRun report for a slide paragraph without creating them.
    private static PowerPointParagraphContent ProjectParagraph(A.Paragraph paragraph, SlidePart? slidePart) {
        var inlines = new List<PowerPointInlineContent>();
        foreach (OpenXmlElement child in paragraph.ChildElements) {
            if (child is A.Run run) {
                A.HyperlinkOnClick? click = run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>();
                Uri? hyperlink = click == null || slidePart == null ? null : PowerPointHyperlinkResolver.Resolve(slidePart, slidePart, click);
                inlines.Add(new PowerPointInlineContent(PowerPointParagraphInlineKind.Run, run.Text?.Text ?? string.Empty, hyperlink, null));
            } else if (child is A.Break) {
                inlines.Add(PowerPointInlineContent.LineBreak);
            } else if (child is A.Field field) {
                inlines.Add(new PowerPointInlineContent(PowerPointParagraphInlineKind.Field, field.Text?.Text ?? string.Empty,
                    null, field.Text == null ? field.InnerText ?? string.Empty : null));
            }
        }
        A.ParagraphProperties? properties = paragraph.ParagraphProperties;
        A.AutoNumberedBullet? numbering = properties?.GetFirstChild<A.AutoNumberedBullet>();
        return new PowerPointParagraphContent(inlines, properties?.Level?.Value,
            properties?.GetFirstChild<A.CharacterBullet>()?.Char?.Value, numbering != null,
            PowerPointSlideReader.ParseNumberingScheme(numbering?.Type?.InnerText), numbering?.StartAt?.Value);
    }
}
