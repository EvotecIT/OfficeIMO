using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class Text {
        public static DocumentFormat.OpenXml.Wordprocessing.Paragraph Add(WordprocessingDocument wordDocument, DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph, string text, int? fontSize = null, bool bold = false, Color color = null, SpaceProcessingModeValues? space = null,
            bool italic = false, DocumentFormat.OpenXml.Wordprocessing.Underline underline = null, bool noProof = false, Spacing spacing = null, TextCapsValues? capital = null,
            bool? kumimoji = null, int? kerning = null, bool? normalizeHeight = null, bool strike = false, bool doubleStrike = false,
            bool caps = false, bool highLight = false) {
            DocumentFormat.OpenXml.Wordprocessing.Text textProp = new DocumentFormat.OpenXml.Wordprocessing.Text();
            textProp.Text = text;
            if (space != null) textProp.Space = space;

            DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties = new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
            if (bold) runProperties.Bold = new Bold();
            if (italic) runProperties.Italic = new Italic();
            if (underline != null) runProperties.Underline = underline;
            if (noProof) runProperties.NoProof = new NoProof();
            //runProperties.AlternativeLanguage = alternativeLanguage;
            if (spacing != null) runProperties.Spacing = spacing;
            //if (capital != null) runProperties.Capital = capital;
            // if (kumimoji != null) runProperties.Kumimoji = kumimoji;
            // if (kerning != null) runProperties.Kerning = kerning;
            // if (normalizeHeight != null) runProperties.NormalizeHeight = normalizeHeight;
            if (strike) runProperties.Strike = new Strike();
            if (fontSize != null) {
                var fontSizeValue = new FontSize {
                    Val = (fontSize * 2).ToString()
                };
                runProperties.FontSize = fontSizeValue;
            }

            if (color != null) {
                runProperties.Color = color;
            }

            //runProperties.Border = new Border();
            if (doubleStrike) {
                runProperties.DoubleStrike = new DoubleStrike();
            }

            if (caps) {
                runProperties.Caps = new Caps();
            }

            if (highLight) {
                runProperties.Highlight = new DocumentFormat.OpenXml.Wordprocessing.Highlight();
            }

            DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();
            run.AppendChild(runProperties);
            run.AppendChild(textProp);

            if (paragraph == null) {
                paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
            }

            paragraph.Append(run);


            return paragraph;
        }
    }
}