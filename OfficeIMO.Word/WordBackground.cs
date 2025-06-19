using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordBackground {

        internal WordDocument _document;

        /// <summary>
        /// Gets or sets the Color.
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

        public WordBackground(WordDocument document) {
            _document = document;

            _document.Background = this;
        }

        public WordBackground(WordDocument document, SixLabors.ImageSharp.Color color) {
            _document = document;

            DocumentBackground documentBackground = new DocumentBackground() { Color = color.ToHexColor() };

            document._document.Body.Append(documentBackground);

            //DocumentBackground documentBackground2 = new DocumentBackground() { Color = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
        }
        /// <summary>
        /// Executes the SetColorHex operation.
        /// </summary>
        public WordBackground SetColorHex(string color) {
            this.Color = color.Replace("#", ""); ;
            return this;
        }
        /// <summary>
        /// Executes the SetColor operation.
        /// </summary>
        public WordBackground SetColor(SixLabors.ImageSharp.Color color) {
            this.Color = color.ToHexColor();
            return this;
        }
    }
}
