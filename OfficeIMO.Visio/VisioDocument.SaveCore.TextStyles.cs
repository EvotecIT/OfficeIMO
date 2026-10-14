using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private void WriteTextBlockCells(XmlWriter writer, string ns, VisioTextStyle? textStyle, bool includeTextTransform = true) {
            if (textStyle == null) {
                return;
            }

            if (textStyle.LeftMargin.HasValue) {
                WriteCell(writer, ns, "LeftMargin", textStyle.LeftMargin.Value);
            }

            if (textStyle.RightMargin.HasValue) {
                WriteCell(writer, ns, "RightMargin", textStyle.RightMargin.Value);
            }

            if (textStyle.TopMargin.HasValue) {
                WriteCell(writer, ns, "TopMargin", textStyle.TopMargin.Value);
            }

            if (textStyle.BottomMargin.HasValue) {
                WriteCell(writer, ns, "BottomMargin", textStyle.BottomMargin.Value);
            }

            if (textStyle.VerticalAlignment.HasValue) {
                WriteCell(writer, ns, "VerticalAlign", (int)textStyle.VerticalAlignment.Value);
            }

            WriteTextBackgroundColorCell(writer, ns, textStyle);
            WriteTextBackgroundTransparencyCell(writer, ns, textStyle);

            if (!includeTextTransform) {
                return;
            }

            if (textStyle.TextPinX.HasValue) {
                WriteCell(writer, ns, "TxtPinX", textStyle.TextPinX.Value);
            }

            if (textStyle.TextPinY.HasValue) {
                WriteCell(writer, ns, "TxtPinY", textStyle.TextPinY.Value);
            }

            if (textStyle.TextWidth.HasValue) {
                WriteCell(writer, ns, "TxtWidth", textStyle.TextWidth.Value);
            }

            if (textStyle.TextHeight.HasValue) {
                WriteCell(writer, ns, "TxtHeight", textStyle.TextHeight.Value);
            }

            if (textStyle.TextLocPinX.HasValue) {
                WriteCell(writer, ns, "TxtLocPinX", textStyle.TextLocPinX.Value);
            }

            if (textStyle.TextLocPinY.HasValue) {
                WriteCell(writer, ns, "TxtLocPinY", textStyle.TextLocPinY.Value);
            }

            if (textStyle.TextAngle.HasValue) {
                WriteCell(writer, ns, "TxtAngle", textStyle.TextAngle.Value);
            }
        }

        private static void WriteTextStyleSections(XmlWriter writer, string ns, VisioTextStyle? textStyle,
            VisioTextSectionSource? character = null, VisioTextSectionSource? paragraph = null,
            IEnumerable<XElement>? nativeSections = null) {
            WriteCharSection(writer, ns, textStyle, character, nativeSections);
            WriteParaSection(writer, ns, textStyle, paragraph, nativeSections);
        }

        private static void WriteCharSection(XmlWriter writer, string ns, VisioTextStyle? textStyle,
            VisioTextSectionSource? source = null, IEnumerable<XElement>? nativeSections = null) {
            // A complex native row set owns its formatting. A separate typed row
            // would duplicate row identities and corrupt the saved document.
            if (nativeSections?.Any(section => IsCharacterSection((string?)section.Attribute("N"))) == true) return;
            if (source != null) {
                WriteTextSectionSource(writer, ns, textStyle, source, character: true);
                return;
            }
            if (!HasCharFormatting(textStyle)) {
                return;
            }

            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Character");
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("IX", "0");

            if (textStyle!.FontFaceId.HasValue) {
                WriteCell(writer, ns, "Font", textStyle.FontFaceId.Value);
            }

            if (textStyle.Color.HasValue) {
                WriteCellValue(writer, ns, "Color", textStyle.Color.Value.ToVisioHex());
            }

            if (textStyle.Size.HasValue) {
                WriteCell(writer, ns, "Size", textStyle.Size.Value / 72D, "PT", null);
            }

            if (TryGetCharStyleValue(textStyle, out int styleValue)) {
                WriteCell(writer, ns, "Style", styleValue);
            }

            if (textStyle.UnderlineStyle.HasValue) {
                WriteCell(writer, ns, "DblUnderline", textStyle.UnderlineStyle.Value == OfficeIMO.Drawing.OfficeTextDecorationStyle.Double ? 1 : 0);
            }

            if (textStyle.StrikethroughStyle.HasValue) {
                WriteCell(writer, ns, "Strikethru", textStyle.StrikethroughStyle.Value == OfficeIMO.Drawing.OfficeTextDecorationStyle.Single ? 1 : 0);
                WriteCell(writer, ns, "DoubleStrikethrough", textStyle.StrikethroughStyle.Value == OfficeIMO.Drawing.OfficeTextDecorationStyle.Double ? 1 : 0);
            }

            if (textStyle.Capitalization.HasValue) {
                WriteCell(writer, ns, "Case", (int)textStyle.Capitalization.Value);
            }

            if (textStyle.Baseline.HasValue) {
                WriteCell(writer, ns, "Pos", (int)textStyle.Baseline.Value);
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteParaSection(XmlWriter writer, string ns, VisioTextStyle? textStyle,
            VisioTextSectionSource? source = null, IEnumerable<XElement>? nativeSections = null) {
            if (nativeSections?.Any(section => IsParagraphSection((string?)section.Attribute("N"))) == true) return;
            if (source != null) {
                WriteTextSectionSource(writer, ns, textStyle, source, character: false);
                return;
            }
            if (textStyle?.HorizontalAlignment == null) {
                return;
            }

            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Paragraph");
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("IX", "0");
            WriteCell(writer, ns, "HorzAlign", (int)textStyle.HorizontalAlignment.Value);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static bool HasCharFormatting(VisioTextStyle? textStyle) {
            return textStyle != null &&
                   (textStyle.FontFaceId.HasValue ||
                    !string.IsNullOrWhiteSpace(textStyle.FontFamily) ||
                    textStyle.Color.HasValue ||
                    textStyle.Size.HasValue ||
                    textStyle.Bold.HasValue ||
                    textStyle.Italic.HasValue ||
                    textStyle.UnderlineStyle.HasValue ||
                    textStyle.StrikethroughStyle.HasValue ||
                    textStyle.SmallCaps.HasValue ||
                    textStyle.Capitalization.HasValue ||
                    textStyle.Baseline.HasValue);
        }

        private static bool TryGetCharStyleValue(VisioTextStyle textStyle, out int styleValue) {
            bool hasAny = false;
            styleValue = 0;
            if (textStyle.Bold.HasValue) {
                hasAny = true;
                if (textStyle.Bold.Value) {
                    styleValue |= 1;
                }
            }

            if (textStyle.Italic.HasValue) {
                hasAny = true;
                if (textStyle.Italic.Value) {
                    styleValue |= 2;
                }
            }

            if (textStyle.UnderlineStyle.HasValue) {
                hasAny = true;
                if (textStyle.UnderlineStyle.Value == OfficeIMO.Drawing.OfficeTextDecorationStyle.Single) {
                    styleValue |= 4;
                }
            }

            if (textStyle.SmallCaps.HasValue) {
                hasAny = true;
                if (textStyle.SmallCaps.Value) {
                    styleValue |= 8;
                }
            }

            return hasAny;
        }
    }
}
