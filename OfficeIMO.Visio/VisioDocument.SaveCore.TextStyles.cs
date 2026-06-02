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

        private static void WriteTextBlockCells(XmlWriter writer, string ns, VisioTextStyle? textStyle, bool includeTextTransform = true) {
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

            if (textStyle.BackgroundColor.HasValue) {
                WriteCellValue(writer, ns, "TextBkgnd", textStyle.BackgroundColor.Value.ToVisioHex());
            }

            if (textStyle.BackgroundTransparency.HasValue) {
                WriteCell(writer, ns, "TextBkgndTrans", textStyle.BackgroundTransparency.Value);
            }

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

        private static void WriteTextStyleSections(XmlWriter writer, string ns, VisioTextStyle? textStyle) {
            WriteCharSection(writer, ns, textStyle);
            WriteParaSection(writer, ns, textStyle);
        }

        private static void WriteCharSection(XmlWriter writer, string ns, VisioTextStyle? textStyle) {
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

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteParaSection(XmlWriter writer, string ns, VisioTextStyle? textStyle) {
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
                    textStyle.Underline.HasValue);
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

            if (textStyle.Underline.HasValue) {
                hasAny = true;
                if (textStyle.Underline.Value) {
                    styleValue |= 4;
                }
            }

            return hasAny;
        }
    }
}
