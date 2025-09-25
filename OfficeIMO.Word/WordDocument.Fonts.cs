using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Embeds a font into the document.
        /// </summary>
        /// <param name="fontPath">Path to a TrueType/OpenType font file.</param>
        public void EmbedFont(string fontPath) {
            if (string.IsNullOrEmpty(fontPath)) throw new ArgumentNullException(nameof(fontPath));
            if (!File.Exists(fontPath)) throw new FileNotFoundException($"Font file '{fontPath}' doesn't exist.", fontPath);

            var mainPart = _wordprocessingDocument?.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            var fontTablePart = mainPart.FontTablePart ?? mainPart.AddNewPart<FontTablePart>();
            fontTablePart.Fonts ??= new Fonts();

            var fontName = Path.GetFileNameWithoutExtension(fontPath);
            var fontPart = fontTablePart.AddFontPart(FontPartType.FontTtf);
            using (var fs = File.OpenRead(fontPath)) {
                fontPart.FeedData(fs);
            }
            var relId = fontTablePart.GetIdOfPart(fontPart);

            var font = fontTablePart.Fonts.Elements<Font>().FirstOrDefault(f => string.Equals(f.Name?.Value, fontName, StringComparison.OrdinalIgnoreCase));
            if (font == null) {
                font = new Font { Name = fontName };
                fontTablePart.Fonts.Append(font);
            }

            font.EmbedRegularFont = new EmbedRegularFont { Id = relId };
        }

        /// <summary>
        /// Embeds a font and registers a paragraph style using that font.
        /// </summary>
        /// <param name="fontPath">Path to a TrueType/OpenType font file.</param>
        /// <param name="styleId">Style identifier to register.</param>
        /// <param name="styleName">Optional friendly name for the style.</param>
        public void EmbedFont(string fontPath, string styleId, string? styleName = null) {
            EmbedFont(fontPath);
            var fontName = Path.GetFileNameWithoutExtension(fontPath);
            WordParagraphStyle.RegisterFontStyle(styleId, fontName, styleName);
            var stylePart = _wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart;
            if (stylePart != null) {
                AddStyleDefinitions(stylePart, overrideExisting: false);
            }
        }
    }
}
