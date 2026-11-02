namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTextRun {
        /// <summary>Sets <see cref="Bold"/> and returns this run.</summary>
        public PowerPointTextRun SetBold(bool isBold = true) {
            Bold = isBold;
            return this;
        }

        /// <summary>Sets <see cref="Italic"/> and returns this run.</summary>
        public PowerPointTextRun SetItalic(bool isItalic = true) {
            Italic = isItalic;
            return this;
        }

        /// <summary>Sets <see cref="Underline"/> and returns this run.</summary>
        public PowerPointTextRun SetUnderline(bool underline = true) {
            Underline = underline;
            return this;
        }

        /// <summary>Sets <see cref="Strikethrough"/> and returns this run.</summary>
        public PowerPointTextRun SetStrikethrough(bool strikethrough = true) {
            Strikethrough = strikethrough;
            return this;
        }

        /// <summary>Sets <see cref="FontSize"/> and returns this run.</summary>
        public PowerPointTextRun SetFontSize(global::OfficeIMO.OfficeLength size) {
            FontSize = size;
            return this;
        }

        /// <summary>Sets <see cref="FontName"/> and returns this run.</summary>
        public PowerPointTextRun SetFontName(string fontName) {
            FontName = fontName;
            return this;
        }

        /// <summary>Sets <see cref="Color"/> and returns this run.</summary>
        public PowerPointTextRun SetColor(string color) {
            Color = color;
            return this;
        }

        /// <summary>Sets <see cref="HighlightColor"/> and returns this run.</summary>
        public PowerPointTextRun SetHighlightColor(string color) {
            HighlightColor = color;
            return this;
        }
    }
}
