using AngleSharp.Css.Dom;
using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        // CSS that decides what content exists and where it breaks, rather than how text looks.
        private static readonly HashSet<string> _structuralCssProperties = new(StringComparer.OrdinalIgnoreCase) {
            "display", "visibility", "white-space", "direction", "unicode-bidi",
            "page-break-before", "page-break-after", "page-break-inside",
            "break-before", "break-after", "break-inside",
            "list-style", "list-style-type", "list-style-position", "counter-reset", "counter-increment",
            "width", "height", "min-width", "max-width", "min-height", "max-height"
        };

        private readonly Dictionary<string, string?> _documentParagraphStyleIds = new(StringComparer.OrdinalIgnoreCase);
        private string? _defaultParagraphStyleId;

        /// <summary>Resolves the document styles the options name, before the conversion changes the document.</summary>
        private void PrepareDocumentStyles(WordDocument doc, HtmlWordExportOptions options) {
            _documentParagraphStyleIds.Clear();
            _defaultParagraphStyleId = null;
            if (string.IsNullOrWhiteSpace(options.DefaultParagraphStyleId)) return;

            _defaultParagraphStyleId = doc.Styles.ResolveParagraphStyleForContent(options.DefaultParagraphStyleId!.Trim())
                ?? throw new InvalidOperationException(
                    $"The default paragraph style '{options.DefaultParagraphStyleId}' is not a paragraph style of the target document " +
                    "or an OfficeIMO built-in paragraph style. Add the style to the document or change DefaultParagraphStyleId.");
        }

        /// <summary>Gives generated body paragraphs without a style of their own the configured default paragraph style.</summary>
        private void ApplyDefaultParagraphStyle(WordDocument doc, IEnumerable<OpenXmlElement> generatedBlocks) {
            if (_defaultParagraphStyleId == null) return;
            foreach (Paragraph paragraph in generatedBlocks.OfType<Paragraph>()) {
                ParagraphProperties? properties = paragraph.ParagraphProperties;
                if (properties?.ParagraphStyleId != null || properties?.NumberingProperties != null) continue;
                new WordParagraph(doc, paragraph).SetStyleId(_defaultParagraphStyleId);
            }
        }

        /// <summary>Finds a paragraph style the target document defines for a CSS class name.</summary>
        private string? FindDocumentParagraphStyle(WordDocument doc, string className) {
            if (!_documentParagraphStyleIds.TryGetValue(className, out string? styleId)) {
                styleId = doc.Styles.FindParagraphStyleId(className);
                _documentParagraphStyleIds[className] = styleId;
            }
            return styleId;
        }

        /// <summary>The option font written to runs; ignored with the document's styles in charge of formatting.</summary>
        private static string? GetRunFontFamilyOption(HtmlWordExportOptions options) =>
            options.IgnoreCssFormatting || string.IsNullOrEmpty(options.FontFamily) ? null : options.FontFamily;

        /// <summary>Drops formatting declarations from an element's style so only structural CSS reaches the importer.</summary>
        private void RemoveCssFormatting(IElement element) {
            string? style = element.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(style)) return;

            ICssStyleDeclaration declaration = ParseInlineDeclaration(style);
            var kept = new StringBuilder();
            foreach (ICssProperty property in declaration) {
                if (!_structuralCssProperties.Contains(property.Name) || string.IsNullOrWhiteSpace(property.Value)) continue;
                kept.Append(property.Name).Append(':').Append(property.Value);
                if (property.IsImportant) kept.Append(" !important");
                kept.Append(';');
            }
            if (kept.Length == 0) {
                element.RemoveAttribute("style");
            } else {
                element.SetAttribute("style", kept.ToString());
            }
        }

        private static void RemoveCssFormatting<TValue>(Dictionary<string, TValue> declarations) {
            foreach (string name in declarations.Keys.Where(name => !_structuralCssProperties.Contains(name)).ToList()) {
                declarations.Remove(name);
            }
        }
    }
}
