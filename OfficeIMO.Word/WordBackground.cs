using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the document background settings.
    /// </summary>
    public class WordBackground {

        internal WordDocument _document;

        /// <summary>
        /// Gets or sets the background color as a hex string.
        /// </summary>
        public string Color {
            get {
                if (_document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground != null) {
                    return _document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground.Color;
                }

                return null;
            }
            set {
                _document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.DisplayBackgroundShape = new DisplayBackgroundShape();
                if (_document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground == null) {
                    _document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground = new DocumentBackground();
                }
                _document._wordprocessingDocument.MainDocumentPart.Document.DocumentBackground.Color = value;
            }
        }

        /// <summary>
        /// Creates a background manager for the specified document.
        /// </summary>
        /// <param name="document">Parent document.</param>
        public WordBackground(WordDocument document) {
            _document = document;

            _document.Background = this;
        }

        /// <summary>
        /// Creates a background with the specified color.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="color">Initial color.</param>
        public WordBackground(WordDocument document, SixLabors.ImageSharp.Color color) {
            _document = document;

            DocumentBackground documentBackground = new DocumentBackground() { Color = color.ToHexColor() };

            document._document.Body.Append(documentBackground);

            //DocumentBackground documentBackground2 = new DocumentBackground() { Color = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
        }
        /// <summary>
        /// Sets the background color using a hex value.
        /// </summary>
        /// <param name="color">Hex color value.</param>
        /// <returns>The current instance.</returns>
        public WordBackground SetColorHex(string color) {
            this.Color = color.Replace("#", ""); ;
            return this;
        }
        /// <summary>
        /// Sets the background color using a <see cref="SixLabors.ImageSharp.Color"/>.
        /// </summary>
        /// <param name="color">Color value.</param>
        /// <returns>The current instance.</returns>
        public WordBackground SetColor(SixLabors.ImageSharp.Color color) {
            this.Color = color.ToHexColor();
            return this;
        }
    }
}
