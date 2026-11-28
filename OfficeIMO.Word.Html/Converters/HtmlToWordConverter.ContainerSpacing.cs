using AngleSharp.Dom;
using OfficeIMO.OpenXml.Internal;
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
                    // Container margins add to the paragraph's own indent, including one its style gives.
                    paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points((paragraph._document.GetEffectiveIndentationBeforeTwips(paragraph) + horizontalStart) / 20D);
                }
                if (horizontalEnd != 0) {
                    paragraph.IndentationAfter = global::OfficeIMO.OfficeLength.Points(((checked((int?)(paragraph.IndentationAfter?.ToTwips())) ?? 0) + horizontalEnd) / 20D);
                }
            }

            foreach (WordTable table in tables) {
                WordTableStyleDetails? styleDetails = table.StyleDetails;
                if (horizontalStart != 0 && styleDetails != null) {
                    long indentation = (styleDetails.TableIndentation?.ToTwips() ?? 0) + horizontalStart;
                    indentation = Math.Max(int.MinValue, Math.Min(int.MaxValue, indentation));
                    styleDetails.TableIndentation = global::OfficeIMO.OfficeLength.Points(indentation / 20D);
                }
                ApplyTableContainerWidth(table, horizontalStart, horizontalEnd);
            }

            ApplyContainerVerticalSpacingFromCss(element, paragraphs, tables);
        }

        private void ApplyTableContainerWidth(
            WordTable table,
            int horizontalStart,
            int horizontalEnd) {
            if (horizontalStart == 0 && horizontalEnd == 0) {
                return;
            }

            int accumulatedSpacing = horizontalStart + horizontalEnd;
            if (_tableContainerHorizontalSpacing.TryGetValue(table.Node(), out int priorSpacing)) {
                accumulatedSpacing += priorSpacing;
            }
            _tableContainerHorizontalSpacing[table.Node()] = accumulatedSpacing;

            int availableWidth = table.EstimateAvailableContainerWidthInDxa();
            long constrainedWidth = (long)availableWidth - accumulatedSpacing;
            if (constrainedWidth <= 0) {
                constrainedWidth = 1;
            } else if (constrainedWidth > int.MaxValue) {
                constrainedWidth = int.MaxValue;
            }

            int currentWidth = table.EstimateTableWidthInDxa();
            if (currentWidth <= constrainedWidth) {
                return;
            }

            table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus((long)constrainedWidth * 635));
        }

        private void ApplyContainerVerticalSpacingFromCss(
            IElement element,
            IReadOnlyList<WordParagraph> paragraphs,
            IReadOnlyList<WordTable> tables) {
            CssStyleMapper.CssProperties box = ParseElementBoxStyles(element);
            OfficeXmlNode first = GetGeneratedBlockBoundary(paragraphs, tables, first: true);
            int verticalStart = (box.MarginTop ?? 0) + (box.PaddingTop ?? 0);
            if (verticalStart != 0 && first.Exists) {
                WordParagraph? paragraph = paragraphs.FirstOrDefault(
                    candidate => candidate.Node() == first);
                if (paragraph != null) {
                    paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(((checked((int?)(paragraph.LineSpacingBefore?.ToTwips())) ?? 0) + verticalStart) / 20D);
                } else {
                    WordTable? table = tables.FirstOrDefault(
                        candidate => candidate.Node() == first);
                    if (table != null) {
                        ApplySpacingAdjacentToTable(table, verticalStart, before: true);
                    }
                }
            }

            OfficeXmlNode last = GetGeneratedBlockBoundary(paragraphs, tables, first: false);
            int verticalEnd = (box.MarginBottom ?? 0) + (box.PaddingBottom ?? 0);
            if (verticalEnd != 0 && last.Exists) {
                WordParagraph? paragraph = paragraphs.FirstOrDefault(
                    candidate => candidate.Node() == last);
                if (paragraph != null) {
                    paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(((checked((int?)(paragraph.LineSpacingAfter?.ToTwips())) ?? 0) + verticalEnd) / 20D);
                } else {
                    WordTable? table = tables.FirstOrDefault(
                        candidate => candidate.Node() == last);
                    if (table != null) {
                        ApplySpacingAdjacentToTable(table, verticalEnd, before: false);
                    }
                }
            }
        }

        private static void ApplySpacingAdjacentToTable(WordTable table, int spacing, bool before) {
            OfficeXmlNode tableNode = table.Node();
            OfficeXmlNode paragraph = before ? default : tableNode.NextSiblingNamed(WordXml.W, "p");
            if (paragraph.IsNone || !HasOnlySyntheticZeroSpacing(paragraph)) {
                OfficeXmlNode created = WordConverterNodes.CreateDetachedParagraph();
                OfficeXmlNode createdProperties = created.Append(created.CreateElement(WordXml.W, "pPr"));
                OfficeXmlNode createdSpacing = createdProperties.Append(createdProperties.CreateElement(WordXml.W, "spacing"));
                createdSpacing.SetAttribute(WordXml.W, "before", "0");
                createdSpacing.SetAttribute(WordXml.W, "after", "0");
                createdSpacing.SetAttribute(WordXml.W, "line", "0");
                paragraph = before
                    ? tableNode.Parent.InsertBefore(created, tableNode)
                    : tableNode.InsertAfterSelf(created);
            }

            OfficeXmlNode properties = paragraph.WChild("pPr");
            if (properties.IsNone) properties = paragraph.InsertBefore(paragraph.CreateElement(WordXml.W, "pPr"), paragraph.FirstChild);
            OfficeXmlNode spacingElement = properties.WChild("spacing");
            if (spacingElement.IsNone) spacingElement = properties.Append(properties.CreateElement(WordXml.W, "spacing"));
            if (before) {
                spacingElement.SetAttribute(WordXml.W, "after", spacing.ToString(CultureInfo.InvariantCulture));
            } else {
                spacingElement.SetAttribute(WordXml.W, "before", spacing.ToString(CultureInfo.InvariantCulture));
            }
            if (spacingElement.WAttr("line") == null) spacingElement.SetAttribute(WordXml.W, "line", "0");
        }

        private static bool HasOnlySyntheticZeroSpacing(OfficeXmlNode paragraph) {
            OfficeXmlNode properties = paragraph.WChild("pPr");
            OfficeXmlNode spacing = properties.WChild("spacing");
            return string.IsNullOrEmpty(paragraph.InnerText) &&
                   WordConverterNodes.ChildNodesIncludingMarkup(paragraph).All(element => element.IsW("pPr")) &&
                   properties.Exists && WordConverterNodes.ChildNodesIncludingMarkup(properties).Count == 1 &&
                   spacing.WAttr("before") == "0" &&
                   spacing.WAttr("after") == "0" &&
                   spacing.WAttr("line") == "0";
        }
    }
}
