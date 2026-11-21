using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Embeds a font into the document.
        /// </summary>
        /// <param name="fontPath">Path to a TrueType/OpenType font file.</param>
        /// <remarks>
        /// The font's OS/2 fsType embedding permission is checked first. Fonts whose licence forbids embedding (Restricted
        /// License without Preview &amp; Print or Editable) or permits only bitmap embedding are refused; Preview &amp; Print,
        /// Editable and Installable fonts are embedded. In a collection, every face must allow embedding.
        /// </remarks>
        /// <exception cref="ArgumentException">The font's embedding permission forbids embedding it.</exception>
        public void EmbedFont(string fontPath) {
            if (string.IsNullOrEmpty(fontPath)) throw new ArgumentNullException(nameof(fontPath));
            if (!File.Exists(fontPath)) throw new FileNotFoundException($"Font file '{fontPath}' doesn't exist.", fontPath);
            byte[] fontData = File.ReadAllBytes(fontPath);
            OfficeIMO.Drawing.OfficeFontEmbeddingPermissions.EnsureFileEmbeddingAllowed(fontData, nameof(fontPath));

            var mainPart = _wordprocessingDocument?.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");
            var fontTablePart = mainPart.FontTablePart ?? mainPart.AddNewPart<FontTablePart>();
            fontTablePart.Fonts ??= new Fonts();

            var fontName = Path.GetFileNameWithoutExtension(fontPath);
            var fontPart = fontTablePart.AddFontPart(FontPartType.FontTtf);
            using (var stream = new MemoryStream(fontData, writable: false)) {
                fontPart.FeedData(stream);
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
        /// Embeds a font and adds a paragraph style to this document that uses it.
        /// </summary>
        /// <param name="fontPath">Path to a TrueType/OpenType font file.</param>
        /// <param name="styleId">Identifier of the paragraph style. An existing style with this identifier is updated to use the font.</param>
        /// <param name="styleName">Optional friendly name for a new style.</param>
        /// <returns>The paragraph style that uses the embedded font.</returns>
        /// <exception cref="ArgumentException">The font's embedding permission forbids embedding it, or the style cannot be used.</exception>
        public WordStyle EmbedFont(string fontPath, string styleId, string? styleName = null) {
            ArgumentException.ThrowIfNullOrWhiteSpace(styleId);
            WordStyle? existing = Styles.FindById(styleId);
            if (existing != null && existing.Type != WordStyleType.Paragraph) {
                throw new ArgumentException($"'{styleId}' is a {existing.Type.ToString().ToLowerInvariant()} style, not a paragraph style.", nameof(styleId));
            }
            // Check the new style's name first, so a rejected style leaves the font table unchanged.
            string newName = string.IsNullOrWhiteSpace(styleName) ? styleId : styleName!;
            if (existing == null && Styles.FindByName(newName) != null) {
                throw new ArgumentException($"The document already has a style named '{newName}'.", nameof(styleName));
            }
            EmbedFont(fontPath);
            WordStyle style = existing ?? Styles.AddParagraphStyle(styleId, styleName);
            style.FontFamily = Path.GetFileNameWithoutExtension(fontPath);
            return style;
        }
    }
}
