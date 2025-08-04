
#nullable enable annotations

namespace OfficeIMO.Html {
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Xml.Linq;
using SixLabors.ImageSharp;
using OfficeIMO.Word.Converters;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeIMO.Html {
    /// <summary>
    /// Converts simple HTML fragments into WordprocessingDocument instances.
    /// </summary>
    public class HtmlToWordConverter : IWordConverter {
        /// <inheritdoc />
        public void Convert(Stream input, Stream output, IConversionOptions options) {
            using StreamReader reader = new StreamReader(input);
            string html = reader.ReadToEnd();
            Convert(html, output, options as HtmlToWordOptions);
        }

        /// <summary>
        /// Converts provided HTML string into a DOCX document written to the specified stream.
        /// </summary>
        /// <param name="html">HTML content to convert. It should be a valid XHTML fragment.</param>
        /// <param name="output">Stream where DOCX content will be written.</param>
        /// <param name="options">Conversion options.</param>
        public static void Convert(string html, Stream output, HtmlToWordOptions? options = null) {
            if (html == null) {
                throw new ArgumentNullException(nameof(html));
            }
            if (output == null) {
                throw new ArgumentNullException(nameof(output));
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
                AppendBlockElement(body, element, options, 0, bulletNumberId, orderedNumberId, mainPart);
            }

            mainPart.Document.Save();
        }

        private static void AppendBlockElement(OpenXmlElement parent, XElement element, HtmlToWordOptions options, int level, int bulletNumberId, int orderedNumberId, MainDocumentPart mainPart) {
            switch (element.Name.LocalName.ToLowerInvariant()) {
                case "p":
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    parent.Append(CreateParagraph(element, options, mainPart));
                    break;
                case "ul":
                    foreach (XElement li in element.Elements("li")) {
                        AppendListItem(parent, li, options, level, bulletNumberId, bulletNumberId, orderedNumberId, mainPart);
                    }
                    break;
                case "ol":
                    foreach (XElement li in element.Elements("li")) {
                        AppendListItem(parent, li, options, level, orderedNumberId, bulletNumberId, orderedNumberId, mainPart);
                    }
                    break;
                case "table":
                    parent.Append(CreateTable(element, options, level, bulletNumberId, orderedNumberId, mainPart));
                    break;
                case "img":
                    Paragraph p = new Paragraph();
                    p.Append(CreateImageRun(element, mainPart));
                    parent.Append(p);
                    break;
            }
        }

        private static Paragraph CreateParagraph(XElement element, HtmlToWordOptions options, MainDocumentPart mainPart) {
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
                if (node is XText textNode) {
                    paragraph.Append(CreateRun(textNode.Value, options));
                } else if (node is XElement inlineElement) {
                    if (inlineElement.Name.LocalName.Equals("img", StringComparison.OrdinalIgnoreCase)) {
                        paragraph.Append(CreateImageRun(inlineElement, mainPart));
                    } else {
                        paragraph.Append(CreateRunFromElement(inlineElement, options));
                    }
                }
            }

            return paragraph;
        }

        private static void AppendListItem(OpenXmlElement parent, XElement li, HtmlToWordOptions options, int level, int numId, int bulletNumberId, int orderedNumberId, MainDocumentPart mainPart) {
            Paragraph paragraph = new Paragraph();
            paragraph.ParagraphProperties = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = level },
                    new NumberingId { Val = numId }
                ));

            foreach (XNode node in li.Nodes()) {
                if (node is XText textNode) {
                    paragraph.Append(CreateRun(textNode.Value, options));
                } else if (node is XElement el) {
                    if (el.Name.LocalName.Equals("ul", StringComparison.OrdinalIgnoreCase) || el.Name.LocalName.Equals("ol", StringComparison.OrdinalIgnoreCase)) {
                        // finalize current paragraph and process nested list
                        parent.Append(paragraph);
                        AppendBlockElement(parent, el, options, level + 1, bulletNumberId, orderedNumberId, mainPart);
                        paragraph = new Paragraph(); // prevent re-adding
                    } else if (el.Name.LocalName.Equals("img", StringComparison.OrdinalIgnoreCase)) {
                        paragraph.Append(CreateImageRun(el, mainPart));
                    } else {
                        paragraph.Append(CreateRunFromElement(el, options));
                    }
                }
            }

            if (paragraph.HasChildren) {
                parent.Append(paragraph);
            }
        }

        private static Table CreateTable(XElement element, HtmlToWordOptions options, int level, int bulletNumberId, int orderedNumberId, MainDocumentPart mainPart) {
            Table table = new Table();

            foreach (XElement tr in element.Elements("tr")) {
                TableRow row = new TableRow();
                foreach (XElement cellEl in tr.Elements().Where(e => e.Name.LocalName.Equals("td", StringComparison.OrdinalIgnoreCase) || e.Name.LocalName.Equals("th", StringComparison.OrdinalIgnoreCase))) {
                    TableCell cell = new TableCell();

                    bool hasBlock = false;
                    foreach (XNode node in cellEl.Nodes()) {
                        if (node is XElement blockEl) {
                            AppendBlockElement(cell, blockEl, options, level, bulletNumberId, orderedNumberId, mainPart);
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

                    row.Append(cell);
                }
                table.Append(row);
            }

            return table;
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
            if (!string.IsNullOrEmpty(options.FontFamily)) {
                RunProperties runProperties = run.RunProperties ??= new RunProperties();
                runProperties.RunFonts = new RunFonts {
                    Ascii = options.FontFamily,
                    HighAnsi = options.FontFamily,
                    ComplexScript = options.FontFamily
                };
            }
            return run;
        }

        private static Run CreateImageRun(XElement element, MainDocumentPart mainPart) {
            string? src = element.Attribute("src")?.Value;
            if (string.IsNullOrEmpty(src)) {
                return new Run();
            }

            byte[] bytes = ResolveImageSource(src);
            using Image image = Image.Load(bytes, out var format);
            long cx = (long)(image.Width * 9525L);
            long cy = (long)(image.Height * 9525L);
            string contentType = format.DefaultMimeType;

            ImagePart imagePart = mainPart.AddImagePart(contentType);
            using (MemoryStream ms = new MemoryStream(bytes)) {
                imagePart.FeedData(ms);
            }
            string relationshipId = mainPart.GetIdOfPart(imagePart);

            var inline = new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = 1U, Name = "Picture" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Image" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            ) { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U };

            var drawing = new Drawing(inline);
            return new Run(drawing);
        }

        private static byte[] ResolveImageSource(string src) {
            if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
                int commaIndex = src.IndexOf(',');
                string base64Data = src.Substring(commaIndex + 1);
                return System.Convert.FromBase64String(base64Data);
            }

            if (Uri.TryCreate(src, UriKind.Absolute, out Uri uri)) {
                if (uri.Scheme == Uri.UriSchemeFile) {
                    return File.ReadAllBytes(uri.LocalPath);
                }
                if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps) {
                    using HttpClient client = new HttpClient();
                    return client.GetByteArrayAsync(uri).GetAwaiter().GetResult();
                }
            }

            if (File.Exists(src)) {
                return File.ReadAllBytes(src);
            }

            throw new InvalidOperationException("Unable to resolve image source: " + src);
        }
    }
}