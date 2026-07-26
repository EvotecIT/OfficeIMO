using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ApplyContainerSpacingFromCss(
            IElement element,
            IReadOnlyList<WordParagraph> paragraphs,
            IReadOnlyList<WordTable> tables) {
            if (paragraphs.Count == 0 && tables.Count == 0) {
                return;
            }

            CssStyleMapper.CssProperties box = ParseElementBoxStyles(element);
            int horizontalStart = (box.MarginLeft ?? 0) + (box.PaddingLeft ?? 0);
            int horizontalEnd = (box.MarginRight ?? 0) + (box.PaddingRight ?? 0);
            foreach (WordParagraph paragraph in DistinctPhysicalParagraphs(paragraphs)) {
                if (horizontalStart != 0) {
                    paragraph.IndentationBefore = (paragraph.IndentationBefore ?? 0) + horizontalStart;
                }
                if (horizontalEnd != 0) {
                    paragraph.IndentationAfter = (paragraph.IndentationAfter ?? 0) + horizontalEnd;
                }
            }

            foreach (WordTable table in tables) {
                WordTableStyleDetails? styleDetails = table.StyleDetails;
                if (horizontalStart != 0 && styleDetails != null) {
                    int indentation = (styleDetails.TableIndentationWidth ?? 0) + horizontalStart;
                    indentation = Math.Max(short.MinValue, Math.Min(short.MaxValue, indentation));
                    styleDetails.TableIndentationWidth = checked((short)indentation);
                }
                ApplyTableContainerWidth(table, horizontalStart, horizontalEnd);
            }

            ApplyContainerVerticalSpacingFromCss(element, paragraphs, tables);
        }

        private static void ApplyTableContainerWidth(
            WordTable table,
            int horizontalStart,
            int horizontalEnd) {
            if (horizontalStart == 0 && horizontalEnd == 0) {
                return;
            }

            int availableWidth = table.EstimateAvailableContainerWidthInDxa();
            long constrainedWidth = (long)availableWidth - horizontalStart - horizontalEnd;
            if (constrainedWidth <= 0) {
                constrainedWidth = 1;
            } else if (constrainedWidth > int.MaxValue) {
                constrainedWidth = int.MaxValue;
            }

            int currentWidth = table.EstimateTableWidthInDxa();
            if (currentWidth <= constrainedWidth) {
                return;
            }

            table.WidthType = TableWidthUnitValues.Dxa;
            table.Width = (int)constrainedWidth;
        }

        private void ApplyContainerVerticalSpacingFromCss(
            IElement element,
            IReadOnlyList<WordParagraph> paragraphs,
            IReadOnlyList<WordTable> tables) {
            CssStyleMapper.CssProperties box = ParseElementBoxStyles(element);
            OpenXmlElement? first = GetGeneratedBlockBoundary(paragraphs, tables, first: true);
            int verticalStart = (box.MarginTop ?? 0) + (box.PaddingTop ?? 0);
            if (verticalStart != 0 && first != null) {
                WordParagraph? paragraph = paragraphs.FirstOrDefault(
                    candidate => ReferenceEquals(candidate._paragraph, first));
                if (paragraph != null) {
                    paragraph.LineSpacingBefore = (paragraph.LineSpacingBefore ?? 0) + verticalStart;
                } else {
                    WordTable? table = tables.FirstOrDefault(
                        candidate => ReferenceEquals(candidate._table, first));
                    if (table != null) {
                        ApplySpacingAdjacentToTable(table, verticalStart, before: true);
                    }
                }
            }

            OpenXmlElement? last = GetGeneratedBlockBoundary(paragraphs, tables, first: false);
            int verticalEnd = (box.MarginBottom ?? 0) + (box.PaddingBottom ?? 0);
            if (verticalEnd != 0 && last != null) {
                WordParagraph? paragraph = paragraphs.FirstOrDefault(
                    candidate => ReferenceEquals(candidate._paragraph, last));
                if (paragraph != null) {
                    paragraph.LineSpacingAfter = (paragraph.LineSpacingAfter ?? 0) + verticalEnd;
                } else {
                    WordTable? table = tables.FirstOrDefault(
                        candidate => ReferenceEquals(candidate._table, last));
                    if (table != null) {
                        ApplySpacingAdjacentToTable(table, verticalEnd, before: false);
                    }
                }
            }
        }

        private static void ApplySpacingAdjacentToTable(WordTable table, int spacing, bool before) {
            Paragraph? paragraph = before ? null : table._table.NextSibling<Paragraph>();
            if (paragraph == null || !HasOnlySyntheticZeroSpacing(paragraph)) {
                paragraph = new Paragraph(
                    new ParagraphProperties(
                        new SpacingBetweenLines {
                            Before = "0",
                            After = "0",
                            Line = "0"
                        }));
                if (before) {
                    table._table.InsertBeforeSelf(paragraph);
                } else {
                    table._table.InsertAfterSelf(paragraph);
                }
            }

            ParagraphProperties properties =
                paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());
            SpacingBetweenLines spacingElement =
                properties.GetFirstChild<SpacingBetweenLines>() ??
                properties.AppendChild(new SpacingBetweenLines());
            if (before) {
                spacingElement.After = spacing.ToString(CultureInfo.InvariantCulture);
            } else {
                spacingElement.Before = spacing.ToString(CultureInfo.InvariantCulture);
            }
            spacingElement.Line ??= "0";
        }

        private static bool HasOnlySyntheticZeroSpacing(Paragraph paragraph) {
            ParagraphProperties? properties = paragraph.ParagraphProperties;
            SpacingBetweenLines? spacing = properties?.GetFirstChild<SpacingBetweenLines>();
            return string.IsNullOrEmpty(paragraph.InnerText) &&
                   paragraph.ChildElements.All(element => element is ParagraphProperties) &&
                   properties?.ChildElements.Count == 1 &&
                   spacing?.Before?.Value == "0" &&
                   spacing.After?.Value == "0" &&
                   spacing.Line?.Value == "0";
        }
    }
}
