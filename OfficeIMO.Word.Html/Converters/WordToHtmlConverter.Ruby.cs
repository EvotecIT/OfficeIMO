using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static bool TryCreateRubyNode(IDocument htmlDoc, WordParagraph run, out INode node) {
            node = htmlDoc.CreateTextNode(string.Empty);
            var ruby = run._run?.Elements<Ruby>().FirstOrDefault();
            if (ruby == null) {
                return false;
            }

            var baseText = ruby.RubyBase?.InnerText ?? string.Empty;
            if (string.IsNullOrEmpty(baseText)) {
                return false;
            }

            var rubyText = ruby.RubyContent?.InnerText ?? string.Empty;
            if (string.IsNullOrEmpty(rubyText)) {
                node = htmlDoc.CreateTextNode(baseText);
                return true;
            }

            var rubyElement = CreateOutputElement(htmlDoc, "ruby");
            var baseElement = CreateOutputElement(htmlDoc, "rb");
            baseElement.TextContent = baseText;
            rubyElement.AppendChild(baseElement);

            var annotationElement = CreateOutputElement(htmlDoc, "rt");
            annotationElement.TextContent = rubyText;
            rubyElement.AppendChild(annotationElement);

            node = rubyElement;
            return true;
        }
    }
}
