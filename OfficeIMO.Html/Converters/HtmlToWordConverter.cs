using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Html {
    /// <summary>
    /// Converts simple HTML fragments into WordprocessingDocument instances.
    /// </summary>
    public class HtmlToWordConverter : IWordConverter {
        /// <summary>
        /// Converts provided HTML string into a DOCX document written to the specified stream.
        /// </summary>
        /// <param name="html">HTML content to convert. It should be a valid XHTML fragment.</param>
        /// <param name="output">Stream where DOCX content will be written.</param>
        /// <param name="options">Conversion options.</param>
        /// <param name="cancellationToken">Token used to cancel the operation.</param>
        public static void Convert(string html, Stream output, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (html == null) {
                throw new ConversionException($"{nameof(html)} cannot be null.");
            }
            if (output == null) {
                throw new ConversionException($"{nameof(output)} cannot be null.");
            }

            options ??= new HtmlToWordOptions();

            using WordprocessingDocument document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document, true);
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            Body body = mainPart.Document.Body;

            // add numbering definitions for ordered and unordered lists using shared Word logic
            NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
            Numbering numbering = WordListStyles.CreateDefaultNumberingDefinitions(document, out int bulletNumberId, out int orderedNumberId);
            numberingPart.Numbering = numbering;

            XDocument xdoc = XDocument.Parse("<root>" + html + "</root>");

            foreach (XElement element in xdoc.Root!.Elements()) {
                cancellationToken.ThrowIfCancellationRequested();
                AppendBlockElement(body, element, options, 0, bulletNumberId, orderedNumberId, mainPart, cancellationToken);
            }

            mainPart.Document.Save();
        }

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
            WordParagraphStyles? style = null;
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

            if (style.HasValue) {
                paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = style.Value.ToString() });
            }

            foreach (XNode node in element.Nodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XText textNode) {
                    paragraph.Append(CreateRun(textNode.Value, options));
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
            paragraph.ParagraphProperties = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = level },
                    new NumberingId { Val = numId }
                ));

            foreach (XNode node in li.Nodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XText textNode) {
                    paragraph.Append(CreateRun(textNode.Value, options));
                } else if (node is XElement el) {
                    if (el.Name.LocalName.Equals("ul", StringComparison.OrdinalIgnoreCase) || el.Name.LocalName.Equals("ol", StringComparison.OrdinalIgnoreCase)) {
                        // finalize current paragraph and process nested list
                        parent.Append(paragraph);
                        AppendBlockElement(parent, el, options, level + 1, bulletNumberId, orderedNumberId, mainPart, cancellationToken);
                        paragraph = new Paragraph(); // prevent re-adding
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
                                p.Append(CreateRun(text.Value, options));
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

        private static Run CreateRunFromElement(XElement element, HtmlToWordOptions options) {
            string text = element.Value;
            Run run = CreateRun(text, options);
            RunProperties runProperties = run.RunProperties ??= new RunProperties();

            switch (element.Name.LocalName.ToLowerInvariant()) {
                case "b":
                case "strong":
                    runProperties.Bold = new Bold();
                    break;
                case "i":
                case "em":
                    runProperties.Italic = new Italic();
                    break;
                case "u":
                    runProperties.Underline = new Underline { Val = UnderlineValues.Single };
                    break;
            }

            return run;
        }

        private static Run CreateRun(string text, HtmlToWordOptions options) {
            Run run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            var fontFamily = FontResolver.Resolve(options.FontFamily);
            if (!string.IsNullOrEmpty(fontFamily)) {
                RunProperties runProperties = run.RunProperties ??= new RunProperties();
                runProperties.RunFonts = new RunFonts {
                    Ascii = fontFamily,
                    HighAnsi = fontFamily,
                    ComplexScript = fontFamily
                };
            }
            return run;
        }

        public void Convert(Stream input, Stream output, IConversionOptions options) {
            if (input == null) {
                throw new ConversionException($"{nameof(input)} cannot be null.");
            }
            using StreamReader reader = new StreamReader(
                input,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 1024,
                leaveOpen: true);
            string html = reader.ReadToEnd();
            Convert(html, output, options as HtmlToWordOptions);
        }

        public async Task ConvertAsync(Stream input, Stream output, IConversionOptions options, CancellationToken cancellationToken = default) {
            if (input == null) {
                throw new ConversionException($"{nameof(input)} cannot be null.");
            }
            using StreamReader reader = new StreamReader(
                input,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 1024,
                leaveOpen: true);
            string html;
#if NET8_0_OR_GREATER
            html = await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
            html = await reader.ReadToEndAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
#endif
            Convert(html, output, options as HtmlToWordOptions, cancellationToken);
            await output.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
    }
}