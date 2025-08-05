using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeIMO.Word.Html {
    public partial class HtmlToWordConverter {
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

        private static void AppendTextWithHyperlinks(OpenXmlElement parent, string text, HtmlToWordOptions options, MainDocumentPart mainPart) {
            int lastIndex = 0;
            foreach (Match match in _urlRegex.Matches(text)) {
                if (match.Index > lastIndex) {
                    parent.Append(CreateRun(text.Substring(lastIndex, match.Index - lastIndex), options));
                }

                string url = match.Value;
                HyperlinkRelationship rel = mainPart.AddHyperlinkRelationship(new Uri(url), true);
                Hyperlink link = new Hyperlink(new Run(new Text(url) { Space = SpaceProcessingModeValues.Preserve })) { Id = rel.Id };
                parent.Append(link);

                lastIndex = match.Index + match.Length;
            }

            if (lastIndex < text.Length) {
                parent.Append(CreateRun(text.Substring(lastIndex), options));
            }
        }
    }
}
