using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html.Helpers;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.Html {
    public partial class HtmlToWordConverter {
        private static void AppendBlockElement(OpenXmlElement parent, XElement element, HtmlToWordOptions options, int level, int bulletNumberId, int orderedNumberId, MainDocumentPart mainPart, CancellationToken cancellationToken) {
            switch (element.Name.LocalName.ToLowerInvariant()) {
                case "p":
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    parent.Append(CreateParagraph(element, options, mainPart, cancellationToken));
                    break;
                case "ul":
                    foreach (XElement li in element.Elements("li")) {
                        cancellationToken.ThrowIfCancellationRequested();
                        AppendListItem(parent, li, options, level, bulletNumberId, bulletNumberId, orderedNumberId, mainPart, cancellationToken);
                    }
                    break;
                case "ol":
                    foreach (XElement li in element.Elements("li")) {
                        cancellationToken.ThrowIfCancellationRequested();
                        AppendListItem(parent, li, options, level, orderedNumberId, bulletNumberId, orderedNumberId, mainPart, cancellationToken);
                    }
                    break;
                case "table":
                    parent.Append(CreateTable(element, options, level, bulletNumberId, orderedNumberId, mainPart, cancellationToken));
                    break;
                case "img":
                    string? src = element.Attribute("src")?.Value;
                    if (!string.IsNullOrEmpty(src)) {
                        Paragraph p = new Paragraph();
                        p.Append(ImageEmbedder.CreateImageRun(mainPart, src));
                        parent.Append(p);
                    }
                    break;
            }
        }

        private static Paragraph CreateParagraph(XElement element, HtmlToWordOptions options, MainDocumentPart mainPart, CancellationToken cancellationToken) {
            Paragraph paragraph = new Paragraph();
            WordParagraphStyles? style = CssStyleMapper.MapParagraphStyle(element.Attribute("style")?.Value);
            if (!style.HasValue) {
                switch (element.Name.LocalName.ToLowerInvariant()) {
                case "h1":
                        style = WordParagraphStyles.Heading1;
                        break;
                case "h2":
                        style = WordParagraphStyles.Heading2;
                        break;
                case "h3":
                        style = WordParagraphStyles.Heading3;
                        break;
                case "h4":
                        style = WordParagraphStyles.Heading4;
                        break;
                case "h5":
                        style = WordParagraphStyles.Heading5;
                        break;
                case "h6":
                        style = WordParagraphStyles.Heading6;
                        break;
                }
            }

            if (style.HasValue) {
                paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = style.Value.ToString() });
            }

            foreach (XNode node in element.Nodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XText textNode) {
                    AppendTextWithHyperlinks(paragraph, textNode.Value, options, mainPart);
                } else if (node is XElement inlineElement) {
                    if (inlineElement.Name.LocalName.Equals("img", StringComparison.OrdinalIgnoreCase)) {
                        string? src = inlineElement.Attribute("src")?.Value;
                        if (!string.IsNullOrEmpty(src)) {
                            paragraph.Append(ImageEmbedder.CreateImageRun(mainPart, src));
                        }
                    } else {
                        paragraph.Append(CreateRunFromElement(inlineElement, options));
                    }
                }
            }

            return paragraph;
        }

        private static void AppendListItem(OpenXmlElement parent, XElement li, HtmlToWordOptions options, int level, int numId, int bulletNumberId, int orderedNumberId, MainDocumentPart mainPart, CancellationToken cancellationToken) {
            Paragraph paragraph = new Paragraph();
            WordParagraphStyles? style = CssStyleMapper.MapParagraphStyle(li.Attribute("style")?.Value);
            ParagraphProperties properties = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = level },
                    new NumberingId { Val = numId }
                ));
            if (style.HasValue) {
                properties.ParagraphStyleId = new ParagraphStyleId { Val = style.Value.ToString() };
            }
            paragraph.ParagraphProperties = properties;

            foreach (XNode node in li.Nodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XText textNode) {
                    AppendTextWithHyperlinks(paragraph, textNode.Value, options, mainPart);
                } else if (node is XElement el) {
                    if (el.Name.LocalName.Equals("ul", StringComparison.OrdinalIgnoreCase) || el.Name.LocalName.Equals("ol", StringComparison.OrdinalIgnoreCase)) {
                        parent.Append(paragraph);
                        AppendBlockElement(parent, el, options, level + 1, bulletNumberId, orderedNumberId, mainPart, cancellationToken);
                        paragraph = new Paragraph();
                    } else if (el.Name.LocalName.Equals("img", StringComparison.OrdinalIgnoreCase)) {
                        string? src = el.Attribute("src")?.Value;
                        if (!string.IsNullOrEmpty(src)) {
                            paragraph.Append(ImageEmbedder.CreateImageRun(mainPart, src));
                        }
                    } else {
                        paragraph.Append(CreateRunFromElement(el, options));
                    }
                }
            }

            if (paragraph.HasChildren) {
                parent.Append(paragraph);
            }
        }

        private static Table CreateTable(XElement element, HtmlToWordOptions options, int level, int bulletNumberId, int orderedNumberId, MainDocumentPart mainPart, CancellationToken cancellationToken) {
            List<List<Action<TableCell>>> structure = new();

            foreach (XElement tr in element.Elements("tr")) {
                cancellationToken.ThrowIfCancellationRequested();
                List<Action<TableCell>> row = new();
                foreach (XElement cellEl in tr.Elements().Where(e => e.Name.LocalName.Equals("td", StringComparison.OrdinalIgnoreCase) || e.Name.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase))) {
                    cancellationToken.ThrowIfCancellationRequested();
                    row.Add(cell => {
                        bool hasBlock = false;
                        foreach (XNode node in cellEl.Nodes()) {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (node is XElement blockEl) {
                                AppendBlockElement(cell, blockEl, options, level, bulletNumberId, orderedNumberId, mainPart, cancellationToken);
                                hasBlock = true;
                            } else if (node is XText text) {
                                Paragraph p = new Paragraph();
                                AppendTextWithHyperlinks(p, text.Value, options, mainPart);
                                cell.Append(p);
                                hasBlock = true;
                            }
                        }

                        if (!hasBlock) {
                            cell.Append(new Paragraph());
                        }
                    });
                }
                structure.Add(row);
            }

            return TableBuilder.Build(structure);
        }
    }
}
