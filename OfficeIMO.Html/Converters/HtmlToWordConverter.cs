using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Html {
    /// <summary>
    /// Converts simple HTML fragments into WordprocessingDocument instances.
    /// </summary>
    public static class HtmlToWordConverter {
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

            // Wrap in a root element to allow multiple top-level paragraphs
            XDocument xdoc = XDocument.Parse("<root>" + html + "</root>");

            foreach (XElement paragraphElement in xdoc.Root!.Elements("p")) {
                Paragraph paragraph = new Paragraph();
                foreach (XNode node in paragraphElement.Nodes()) {
                    if (node is XText textNode) {
                        paragraph.Append(CreateRun(textNode.Value, options));
                    } else if (node is XElement element) {
                        paragraph.Append(CreateRunFromElement(element, options));
                    }
                }
                body.Append(paragraph);
            }

            mainPart.Document.Save();
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
    }
}
