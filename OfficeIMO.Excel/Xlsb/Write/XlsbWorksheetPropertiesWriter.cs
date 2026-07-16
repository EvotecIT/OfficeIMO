using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes the supported BrtWsProp worksheet-property subset.</summary>
    internal static class XlsbWorksheetPropertiesWriter {
        private const int BrtWsProp = 147;

        internal static void Write(Stream output, SheetProperties? properties, string sheetName) {
            if (output == null) throw new ArgumentNullException(nameof(output));
            if (properties == null) return;
            XlsbRecordWriter.Write(output, BrtWsProp, CreatePayload(properties, sheetName));
        }

        private static byte[] CreatePayload(SheetProperties properties, string sheetName) {
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

            TabColor? tabColor = null;
            OutlineProperties? outline = null;
            PageSetupProperties? pageSetup = null;
            foreach (OpenXmlElement child in properties.ChildElements) {
                switch (child) {
                    case TabColor value when tabColor == null:
                        tabColor = value;
                        break;
                    case OutlineProperties value when outline == null:
                        outline = value;
                        break;
                    case PageSetupProperties value when pageSetup == null:
                        pageSetup = value;
                        break;
                    default:
                        throw new NotSupportedException($"Native XLSB generation does not support worksheet property '{child.LocalName}' or duplicates it on worksheet '{sheetName}'.");
                }
            }
            if (outline != null) EnsureOnlyAttributes(outline, sheetName, "applyStyles", "summaryBelow", "summaryRight", "showOutlineSymbols");
            if (pageSetup != null) EnsureOnlyAttributes(pageSetup, sheetName, "autoPageBreaks", "fitToPage");
            if (tabColor != null) {
                if (tabColor.HasChildren) {
                    throw new NotSupportedException($"Native XLSB generation does not support child content in the tab color on worksheet '{sheetName}'.");
                }
                EnsureOnlyAttributes(tabColor, sheetName, "auto", "indexed", "rgb", "theme", "tint");
            }
            if (outline?.HasChildren == true || pageSetup?.HasChildren == true) {
                throw new NotSupportedException($"Native XLSB generation does not support nested worksheet-property content on worksheet '{sheetName}'.");
            }

            bool syncHorizontal = properties.SyncHorizontal?.Value == true;
            bool syncVertical = properties.SyncVertical?.Value == true;
            uint synchronizedRow = uint.MaxValue;
            uint synchronizedColumn = uint.MaxValue;
            if (properties.SyncReference?.Value is string syncReference && syncReference.Length != 0) {
                if (!syncHorizontal && !syncVertical) {
                    throw new NotSupportedException($"Native XLSB generation requires syncHorizontal or syncVertical when syncRef is set on worksheet '{sheetName}'.");
                }
                if (!A1.TryParseCellReferenceFast(syncReference, out int row, out int column)
                    || row <= 0 || row > A1.MaxRows || column <= 0 || column > A1.MaxColumns) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode invalid synchronized-scroll reference '{syncReference}' on worksheet '{sheetName}'.");
                }
                synchronizedRow = checked((uint)row - 1U);
                synchronizedColumn = checked((uint)column - 1U);
            }

            string codeName = properties.CodeName?.Value ?? string.Empty;
            if (codeName.Length > 31) {
                throw new NotSupportedException($"Native XLSB generation limits worksheet code names to 31 characters on worksheet '{sheetName}'.");
            }

            uint flags = 0;
            if (pageSetup?.AutoPageBreaks?.Value == true) flags |= 0x000001U;
            if (properties.Published?.Value == true) flags |= 0x000008U;
            if (outline?.ApplyStyles?.Value == true) flags |= 0x000020U;
            if (outline?.SummaryBelow?.Value != false) flags |= 0x000040U;
            if (outline?.SummaryRight?.Value != false) flags |= 0x000080U;
            if (pageSetup?.FitToPage?.Value == true) flags |= 0x000100U;
            if (outline?.ShowOutlineSymbols?.Value != false) flags |= 0x000400U;
            if (syncHorizontal) flags |= 0x001000U;
            if (syncVertical) flags |= 0x002000U;
            if (properties.TransitionEvaluation?.Value == true) flags |= 0x004000U;
            if (properties.TransitionEntry?.Value == true) flags |= 0x008000U;
            if (properties.FilterMode?.Value == true) flags |= 0x010000U;
            if (properties.EnableFormatConditionsCalculation?.Value != false) flags |= 0x020000U;

            using var output = new MemoryStream(27 + codeName.Length * 2);
            output.WriteByte((byte)flags);
            output.WriteByte((byte)(flags >> 8));
            output.WriteByte((byte)(flags >> 16));
            WriteColor(output, tabColor, sheetName);
            WriteUInt32(output, synchronizedRow);
            WriteUInt32(output, synchronizedColumn);
            WriteWideString(output, codeName);
            return output.ToArray();
        }

        private static void WriteColor(Stream output, TabColor? color, string sheetName) {
            byte type = 0;
            byte index = 0;
            byte red = 0;
            byte green = 0;
            byte blue = 0;
            byte alpha = 0;
            int sourceCount = (color?.Rgb?.Value is string rgbValue && rgbValue.Length > 0 ? 1 : 0)
                + (color?.Theme != null ? 1 : 0)
                + (color?.Indexed != null ? 1 : 0)
                + (color?.Auto?.Value == true ? 1 : 0);
            if (sourceCount > 1) {
                throw new NotSupportedException($"Native XLSB generation requires exactly one tab-color source on worksheet '{sheetName}'.");
            }
            if (color?.Rgb?.Value is string rgb && rgb.Length > 0) {
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
            } else if (color?.Theme?.Value is uint theme) {
                if (theme > 11U) throw new NotSupportedException($"Native XLSB generation cannot encode theme tab color {theme} on worksheet '{sheetName}'.");
                type = 3;
                index = (byte)theme;
            } else if (color?.Indexed?.Value is uint indexed) {
                if (indexed > byte.MaxValue) throw new NotSupportedException($"Native XLSB generation cannot encode indexed tab color {indexed} on worksheet '{sheetName}'.");
                type = 1;
                index = (byte)indexed;
            } else if (color?.Auto?.Value == true) {
                type = 0;
            }

            output.WriteByte((byte)((type << 1) | (type == 2 ? 0x01 : 0)));
            output.WriteByte(index);
            WriteUInt16(output, unchecked((ushort)ToTint(color?.Tint?.Value ?? 0D, sheetName)));
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
            WriteUInt32(output, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void WriteUInt16(Stream output, ushort value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
        }

        private static void WriteUInt32(Stream output, uint value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
            output.WriteByte((byte)(value >> 16));
            output.WriteByte((byte)(value >> 24));
        }

        private static void EnsureOnlyAttributes(OpenXmlElement element, string sheetName, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support worksheet-property attribute '{unsupported.Value.LocalName}' on worksheet '{sheetName}'.");
            }
        }
    }
}
