using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// Describes the shapes of a loaded slide with the content records <see cref="PowerPointSlideReader"/> returns, so
/// extraction renders a loaded presentation and a streamed one through the same code.
/// </summary>
internal static class PowerPointSlideContentProjection {
    internal static IEnumerable<PowerPointShapeContent> ProjectShapes(PowerPointSlide slide, bool includeHidden) {
        foreach (PowerPointShape shape in slide.EnumerateShapesDeep(slide.Shapes, includeHidden)) {
            yield return ProjectShape(shape);
        }
    }

    private static PowerPointShapeContent ProjectShape(PowerPointShape shape) {
        IReadOnlyList<PowerPointParagraphContent> paragraphs = shape is PowerPointTextBox textBox
            ? textBox.Paragraphs.Select(ProjectParagraph).ToList()
            : Array.Empty<PowerPointParagraphContent>();
        PowerPointTableContent? table = shape is PowerPointTable powerPointTable ? ProjectTable(powerPointTable) : null;
        return new PowerPointShapeContent(shape.ShapeContentType, shape.Id, shape.Name, shape.AltText, shape.Hidden,
            isVisible: !shape.Hidden, depth: 0, shape.ShapePlaceholderType, paragraphs, table, image: null);
    }

    private static PowerPointTableContent ProjectTable(PowerPointTable table) {
        A.TableGrid? grid = table.TableElement.TableGrid;
        var rows = table.Rows
            .Select(row => new PowerPointTableRowContent(row.Cells
                .Select(cell => new PowerPointTableCellContent(cell.Paragraphs.Select(ProjectParagraph).ToList()))
                .ToList()))
            .ToList();
        return new PowerPointTableContent(table.Name, grid?.Elements<A.GridColumn>().Count() ?? 0, table.HeaderRow, rows,
            hasGrid: grid != null);
    }

    internal static PowerPointParagraphContent ProjectParagraph(PowerPointParagraph paragraph) {
        // Run wrappers resolve hyperlinks through the owning part; build them only when a run has a click link.
        IReadOnlyList<PowerPointTextRun>? runs = null;
        int runIndex = 0;
        var inlines = new List<PowerPointInlineContent>();
        foreach (OpenXmlElement child in paragraph.Paragraph.ChildElements) {
            if (child is A.Run run) {
                Uri? hyperlink = null;
                if (run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>() != null) {
                    runs ??= paragraph.Runs;
                    hyperlink = runs[runIndex].Hyperlink;
                }
                runIndex++;
                inlines.Add(new PowerPointInlineContent(PowerPointParagraphInlineKind.Run, run.Text?.Text ?? string.Empty, hyperlink, null));
            } else if (child is A.Break) {
                inlines.Add(PowerPointInlineContent.LineBreak);
            } else if (child is A.Field field) {
                inlines.Add(new PowerPointInlineContent(PowerPointParagraphInlineKind.Field, field.Text?.Text ?? string.Empty,
                    null, field.Text == null ? field.InnerText ?? string.Empty : null));
            }
        }
        return new PowerPointParagraphContent(inlines, paragraph.Level, paragraph.BulletCharacter, paragraph.IsNumbered,
            paragraph.NumberingScheme, paragraph.NumberingStartAt);
    }
}
