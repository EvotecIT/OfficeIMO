using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordFormatting {
        internal DocumentFormat.OpenXml.Wordprocessing.RunProperties _runProperties = null;
        public bool Bold   // property
        {
            get {
                if (_runProperties.Bold != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                if (value != true) {
                    _runProperties.Bold = null;
                } else {
                    _runProperties.Bold = new Bold();
                }
            }
        }

        public WordFormatting() {
            DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            _runProperties = runProperties;
        }

        public WordFormatting(string test) {
            DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            bool bold = false;
            if (bold) runProperties.Bold = new Bold();
            bool italic = false;
            if (italic) runProperties.Italic = new Italic();
            Underline underline = null;
            if (underline != null) runProperties.Underline = underline;
            bool noProof = false;
            if (noProof) runProperties.NoProof = new NoProof();
            //runProperties.AlternativeLanguage = alternativeLanguage;
            Spacing spacing = null;
            if (spacing != null) runProperties.Spacing = spacing;
            //if (capital != null) runProperties.Capital = capital;
            // if (kumimoji != null) runProperties.Kumimoji = kumimoji;
            // if (kerning != null) runProperties.Kerning = kerning;
            // if (normalizeHeight != null) runProperties.NormalizeHeight = normalizeHeight;
            bool strike = false;
            if (strike) runProperties.Strike = new Strike();
            int? fontSize = null;
            if (fontSize != null) {
                var fontSizeValue = new FontSize {
                    Val = (fontSize * 2).ToString()
                };
                runProperties.FontSize = fontSizeValue;
            }

            Color color = null;
            if (color != null) {
                runProperties.Color = color;
            }

            //runProperties.Border = new Border();
            bool doubleStrike = false;
            if (doubleStrike) {
                runProperties.DoubleStrike = new DoubleStrike();
            }

            bool caps = false;
            if (caps) {
                runProperties.Caps = new Caps();
            }

            bool highLight = false;
            if (highLight) {
                runProperties.Highlight = new DocumentFormat.OpenXml.Wordprocessing.Highlight();
            }

            DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();
            run.AppendChild(runProperties);
        }
    }
}