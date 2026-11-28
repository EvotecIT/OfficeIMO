using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes the supported BrtWsProp worksheet-property subset.</summary>
    internal static class XlsbWorksheetPropertiesWriter {
        private const int BrtWsProp = 147;

        internal static void Write(Stream output, OfficeXmlNode properties, string sheetName, string? codeNameOverride) {
            ArgumentNullException.ThrowIfNull(output);
            if (properties.IsNone && codeNameOverride == null) return;
            XlsbRecordWriter.Write(output, BrtWsProp, CreatePayload(properties, sheetName, codeNameOverride));
        }

        private static byte[] CreatePayload(OfficeXmlNode properties, string sheetName, string? codeNameOverride) {
            EnsureOnlyAttributes(
                properties,
                sheetName,
                "syncHorizontal",
                "syncVertical",
                "syncRef",
                "transitionEvaluation",
                "transitionEntry",
                "published",
                "codeName",
                "filterMode",
                "enableFormatConditionsCalculation");

            OfficeXmlNode tabColor = default, outline = default, pageSetup = default;
            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.NamespaceId == OfficeXmlNamespaces.SpreadsheetML) {
                    switch (child.LocalName) {
                        case "tabColor" when tabColor.IsNone: tabColor = child; continue;
                        case "outlinePr" when outline.IsNone: outline = child; continue;
                        case "pageSetUpPr" when pageSetup.IsNone: pageSetup = child; continue;
                    }
                }
                throw new NotSupportedException($"Native XLSB generation does not support worksheet property '{child.LocalName}' or duplicates it on worksheet '{sheetName}'.");
            }
            if (outline.Exists) EnsureOnlyAttributes(outline, sheetName, "applyStyles", "summaryBelow", "summaryRight", "showOutlineSymbols");
            if (pageSetup.Exists) EnsureOnlyAttributes(pageSetup, sheetName, "autoPageBreaks", "fitToPage");
            if (tabColor.Exists) {
                if (tabColor.FirstChild.Exists) {
                    throw new NotSupportedException($"Native XLSB generation does not support child content in the tab color on worksheet '{sheetName}'.");
                }
                EnsureOnlyAttributes(tabColor, sheetName, "auto", "indexed", "rgb", "theme", "tint");
            }
            if (outline.FirstChild.Exists || pageSetup.FirstChild.Exists) {
                throw new NotSupportedException($"Native XLSB generation does not support nested worksheet-property content on worksheet '{sheetName}'.");
            }

            bool syncHorizontal = XlsbWorkbookMarkup.ReadBoolean(properties, "syncHorizontal") == true;
            bool syncVertical = XlsbWorkbookMarkup.ReadBoolean(properties, "syncVertical") == true;
            uint synchronizedRow = uint.MaxValue;
            uint synchronizedColumn = uint.MaxValue;
            if (properties.Attribute("syncRef") is string syncReference && syncReference.Length != 0) {
                if (!syncHorizontal && !syncVertical) {
                    throw new NotSupportedException($"Native XLSB generation requires syncHorizontal or syncVertical when syncRef is set on worksheet '{sheetName}'.");
                }
                if (!ExcelA1.TryParseCellReferenceFast(syncReference, out int row, out int column)
                    || row <= 0 || row > ExcelA1.MaxRows || column <= 0 || column > ExcelA1.MaxColumns) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode invalid synchronized-scroll reference '{syncReference}' on worksheet '{sheetName}'.");
                }
                synchronizedRow = checked((uint)row - 1U);
                synchronizedColumn = checked((uint)column - 1U);
            }

            string codeName = codeNameOverride ?? properties.Attribute("codeName") ?? string.Empty;
            if (codeName.Length > 31) {
                throw new NotSupportedException($"Native XLSB generation limits worksheet code names to 31 characters on worksheet '{sheetName}'.");
            }

            uint flags = 0;
            if (XlsbWorkbookMarkup.ReadBoolean(pageSetup, "autoPageBreaks") == true) flags |= 0x000001U;
            if (XlsbWorkbookMarkup.ReadBoolean(properties, "published") == true) flags |= 0x000008U;
            if (XlsbWorkbookMarkup.ReadBoolean(outline, "applyStyles") == true) flags |= 0x000020U;
            if (XlsbWorkbookMarkup.ReadBoolean(outline, "summaryBelow") != false) flags |= 0x000040U;
            if (XlsbWorkbookMarkup.ReadBoolean(outline, "summaryRight") != false) flags |= 0x000080U;
            if (XlsbWorkbookMarkup.ReadBoolean(pageSetup, "fitToPage") == true) flags |= 0x000100U;
            if (XlsbWorkbookMarkup.ReadBoolean(outline, "showOutlineSymbols") != false) flags |= 0x000400U;
            if (syncHorizontal) flags |= 0x001000U;
            if (syncVertical) flags |= 0x002000U;
            if (XlsbWorkbookMarkup.ReadBoolean(properties, "transitionEvaluation") == true) flags |= 0x004000U;
            if (XlsbWorkbookMarkup.ReadBoolean(properties, "transitionEntry") == true) flags |= 0x008000U;
            if (XlsbWorkbookMarkup.ReadBoolean(properties, "filterMode") == true) flags |= 0x010000U;
            if (XlsbWorkbookMarkup.ReadBoolean(properties, "enableFormatConditionsCalculation") != false) flags |= 0x020000U;

            using var output = new MemoryStream(27 + codeName.Length * 2);
            output.WriteByte((byte)flags);
            output.WriteByte((byte)(flags >> 8));
            output.WriteByte((byte)(flags >> 16));
            WriteColor(output, tabColor, sheetName);
            XlsbPayloadWriter.WriteUInt32(output, synchronizedRow);
            XlsbPayloadWriter.WriteUInt32(output, synchronizedColumn);
            WriteWideString(output, codeName);
            return output.ToArray();
        }

        private static void WriteColor(Stream output, OfficeXmlNode color, string sheetName) {
            byte type = 0;
            byte index = 0;
            byte red = 0;
            byte green = 0;
            byte blue = 0;
            byte alpha = 0;
            int sourceCount = (color.Attribute("rgb") is string rgbValue && rgbValue.Length > 0 ? 1 : 0)
                + (color.HasAttribute("theme") ? 1 : 0)
                + (color.HasAttribute("indexed") ? 1 : 0)
                + (XlsbWorkbookMarkup.ReadBoolean(color, "auto") == true ? 1 : 0);
            if (sourceCount > 1) {
                throw new NotSupportedException($"Native XLSB generation requires exactly one tab-color source on worksheet '{sheetName}'.");
            }
            if (color.Attribute("rgb") is string rgb && rgb.Length > 0) {
                string normalized = rgb.Length == 6 ? "FF" + rgb : rgb;
                if (normalized.Length != 8
                    || !uint.TryParse(normalized, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint argb)) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode tab color '{rgb}' on worksheet '{sheetName}'.");
                }
                type = 2;
                alpha = (byte)(argb >> 24);
                red = (byte)(argb >> 16);
                green = (byte)(argb >> 8);
                blue = (byte)argb;
            } else if (XlsbWorkbookMarkup.ReadUInt32(color, "theme") is uint theme) {
                if (theme > 11U) throw new NotSupportedException($"Native XLSB generation cannot encode theme tab color {theme} on worksheet '{sheetName}'.");
                type = 3;
                index = (byte)theme;
            } else if (XlsbWorkbookMarkup.ReadUInt32(color, "indexed") is uint indexed) {
                if (indexed > byte.MaxValue) throw new NotSupportedException($"Native XLSB generation cannot encode indexed tab color {indexed} on worksheet '{sheetName}'.");
                type = 1;
                index = (byte)indexed;
            }

            output.WriteByte((byte)((type << 1) | (type == 2 ? 0x01 : 0)));
            output.WriteByte(index);
            XlsbPayloadWriter.WriteUInt16(output, unchecked((ushort)ToTint(color.Attribute("tint") is string tint ? System.Xml.XmlConvert.ToDouble(tint) : 0D, sheetName)));
            output.WriteByte(red);
            output.WriteByte(green);
            output.WriteByte(blue);
            output.WriteByte(alpha);
        }

        private static short ToTint(double value, string sheetName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < -1D || value > 1D) {
                throw new NotSupportedException($"Native XLSB generation cannot encode tab-color tint {value} on worksheet '{sheetName}'.");
            }
            return checked((short)Math.Round(value * (value < 0D ? 32768D : 32767D), MidpointRounding.AwayFromZero));
        }

        private static void WriteWideString(Stream output, string value) {
            XlsbPayloadWriter.WriteUInt32(output, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void EnsureOnlyAttributes(OfficeXmlNode element, string sheetName, params string[] allowedNames) {
            string? unsupported = XlsbWorkbookMarkup.UnsupportedAttribute(element, allowedNames);
            if (unsupported != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support worksheet-property attribute '{unsupported}' on worksheet '{sheetName}'.");
            }
        }
    }
}
