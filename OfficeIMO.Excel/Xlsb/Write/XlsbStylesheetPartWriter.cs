using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Writes the core BIFF12 formatting collections referenced by normal worksheet cells.</summary>
    internal static class XlsbStylesheetPartWriter {
        private const int MaximumStyleNameLength = 255;
        private const int BrtBeginStyleSheet = 278;
        private const int BrtEndStyleSheet = 279;
        private const int BrtBeginFills = 603;
        private const int BrtEndFills = 604;
        private const int BrtBeginFonts = 611;
        private const int BrtEndFonts = 612;
        private const int BrtBeginBorders = 613;
        private const int BrtEndBorders = 614;
        private const int BrtBeginFmts = 615;
        private const int BrtEndFmts = 616;
        private const int BrtBeginCellXfs = 617;
        private const int BrtEndCellXfs = 618;
        private const int BrtBeginStyles = 619;
        private const int BrtEndStyles = 620;
        private const int BrtBeginCellStyleXfs = 626;
        private const int BrtEndCellStyleXfs = 627;
        private const int BrtBeginDxfs = 505;
        private const int BrtEndDxfs = 506;
        private const int BrtBeginTableStyles = 508;
        private const int BrtEndTableStyles = 509;
        private const int BrtFont = 43;
        private const int BrtFmt = 44;
        private const int BrtFill = 45;
        private const int BrtBorder = 46;
        private const int BrtXf = 47;
        private const int BrtStyle = 48;

        internal static byte[] Create(Stylesheet stylesheet, out int cellFormatCount) {
            if (stylesheet == null) throw new ArgumentNullException(nameof(stylesheet));
            ValidateUnsupportedCollections(stylesheet);

            NumberingFormat[] numberFormats = stylesheet.NumberingFormats?.Elements<NumberingFormat>().ToArray()
                ?? Array.Empty<NumberingFormat>();
            Font[] fonts = RequireItems(stylesheet.Fonts?.Elements<Font>(), "fonts");
            Fill[] fills = RequireItems(stylesheet.Fills?.Elements<Fill>(), "fills");
            Border[] borders = RequireItems(stylesheet.Borders?.Elements<Border>(), "borders");
            CellFormat[] styleFormats = RequireItems(stylesheet.CellStyleFormats?.Elements<CellFormat>(), "cell style formats");
            CellFormat[] cellFormats = RequireItems(stylesheet.CellFormats?.Elements<CellFormat>(), "cell formats");
            cellFormatCount = cellFormats.Length;

            ValidateCellFormatReferences(styleFormats, fonts.Length, fills.Length, borders.Length, isStyleFormat: true);
            ValidateCellFormatReferences(cellFormats, fonts.Length, fills.Length, borders.Length, isStyleFormat: false, styleFormats.Length);

            using var output = new MemoryStream(1024 + (fonts.Length + fills.Length + borders.Length + cellFormats.Length) * 80);
            XlsbRecordWriter.Write(output, BrtBeginStyleSheet);
            if (numberFormats.Length > 0) {
                WriteCollection(output, BrtBeginFmts, BrtEndFmts, numberFormats, WriteNumberFormat);
            }
            WriteCollection(output, BrtBeginFonts, BrtEndFonts, fonts, WriteFont);
            WriteCollection(output, BrtBeginFills, BrtEndFills, fills, WriteFill);
            WriteCollection(output, BrtBeginBorders, BrtEndBorders, borders, WriteBorder);
            WriteCollection(output, BrtBeginCellStyleXfs, BrtEndCellStyleXfs, styleFormats,
                (stream, format) => WriteCellFormat(stream, format, isStyleFormat: true));
            WriteCollection(output, BrtBeginCellXfs, BrtEndCellXfs, cellFormats,
                (stream, format) => WriteCellFormat(stream, format, isStyleFormat: false));
            WriteNormalStyle(output, stylesheet.CellStyles, styleFormats.Length);
            WriteUInt32Record(output, BrtBeginDxfs, 0);
            XlsbRecordWriter.Write(output, BrtEndDxfs);
            WriteTableStyles(output, stylesheet.TableStyles);
            XlsbRecordWriter.Write(output, BrtEndStyleSheet);
            return output.ToArray();
        }

        private static void ValidateUnsupportedCollections(Stylesheet stylesheet) {
            if (stylesheet.DifferentialFormats?.Elements<DifferentialFormat>().Any() == true) {
                throw new NotSupportedException("Native XLSB generation does not yet encode differential formatting records.");
            }
            if (stylesheet.Colors != null || stylesheet.ChildElements.Any(element => element.LocalName == "extLst")) {
                throw new NotSupportedException("Native XLSB generation does not yet encode custom indexed colors or style extensions.");
            }

            CellStyle[] styles = stylesheet.CellStyles?.Elements<CellStyle>().ToArray() ?? Array.Empty<CellStyle>();
            if (styles.Length > 1
                || styles.Any(style => (style.BuiltinId?.Value ?? 0U) != 0U || (style.FormatId?.Value ?? 0U) != 0U)) {
                throw new NotSupportedException("Native XLSB generation currently supports only the built-in Normal cell style.");
            }
            if (stylesheet.TableStyles?.ChildElements.Count > 0) {
                throw new NotSupportedException("Native XLSB generation does not yet encode custom table or pivot styles.");
            }
        }

        private static T[] RequireItems<T>(IEnumerable<T>? items, string collectionName) where T : OpenXmlElement {
            T[] values = items?.ToArray() ?? Array.Empty<T>();
            if (values.Length == 0) {
                throw new NotSupportedException($"Native XLSB generation requires a non-empty {collectionName} collection.");
            }
            if (values.Length > 65_536) {
                throw new NotSupportedException($"Native XLSB generation supports at most 65,536 {collectionName}.");
            }
            return values;
        }

        private static void ValidateCellFormatReferences(
            IEnumerable<CellFormat> formats,
            int fontCount,
            int fillCount,
            int borderCount,
            bool isStyleFormat,
            int styleFormatCount = 0) {
            foreach (CellFormat format in formats) {
                uint fontId = format.FontId?.Value ?? 0U;
                uint fillId = format.FillId?.Value ?? 0U;
                uint borderId = format.BorderId?.Value ?? 0U;
                uint parentId = format.FormatId?.Value ?? 0U;
                if (fontId >= fontCount || fillId >= fillCount || borderId >= borderCount) {
                    throw new NotSupportedException("Native XLSB generation found a cell format with an out-of-range font, fill, or border reference.");
                }
                if (!isStyleFormat && parentId >= styleFormatCount) {
                    throw new NotSupportedException("Native XLSB generation found a cell format with an out-of-range parent style reference.");
                }
            }
        }

        private static void WriteCollection<T>(
            Stream output,
            int beginRecord,
            int endRecord,
            IReadOnlyList<T> items,
            Action<Stream, T> writeItem) {
            WriteUInt32Record(output, beginRecord, checked((uint)items.Count));
            foreach (T item in items) writeItem(output, item);
            XlsbRecordWriter.Write(output, endRecord);
        }

        private static void WriteNumberFormat(Stream output, NumberingFormat format) {
            uint id = format.NumberFormatId?.Value
                ?? throw new NotSupportedException("Native XLSB generation requires every custom number format to have an id.");
            string code = format.FormatCode?.Value ?? string.Empty;
            if (id > ushort.MaxValue || string.IsNullOrEmpty(code) || code.Length > 255) {
                throw new NotSupportedException("Native XLSB generation found an invalid custom number format.");
            }

            using var payload = new MemoryStream(8 + code.Length * 2);
            WriteUInt16(payload, (ushort)id);
            WriteWideString(payload, code);
            XlsbRecordWriter.Write(output, BrtFmt, payload.ToArray());
        }

        private static void WriteFont(Stream output, Font font) {
            string name = font.FontName?.Val?.Value ?? "Calibri";
            double size = font.FontSize?.Val?.Value ?? 11D;
            if (string.IsNullOrWhiteSpace(name) || name.Length > 31 || size <= 0D || size > 409.55D) {
                throw new NotSupportedException("Native XLSB generation found a font with an unsupported name or size.");
            }

            ushort flags = 0;
            if (font.Bold != null) flags |= 0x0001;
            if (font.Italic != null) flags |= 0x0002;
            if (font.Strike != null) flags |= 0x0008;
            if (font.Outline != null) flags |= 0x0010;
            if (font.Shadow != null) flags |= 0x0020;
            if (font.Condense != null) flags |= 0x0040;
            if (font.Extend != null) flags |= 0x0080;

            using var payload = new MemoryStream(32 + name.Length * 2);
            WriteUInt16(payload, checked((ushort)Math.Round(size * 20D, MidpointRounding.AwayFromZero)));
            WriteUInt16(payload, flags);
            WriteUInt16(payload, font.Bold != null ? (ushort)700 : (ushort)400);
            WriteUInt16(payload, ToScript(font.VerticalTextAlignment?.Val?.Value));
            payload.WriteByte(ToUnderline(font.Underline?.Val?.Value));
            payload.WriteByte(checked((byte)(font.FontFamilyNumbering?.Val?.Value ?? 0)));
            payload.WriteByte(checked((byte)(font.GetFirstChild<FontCharSet>()?.Val?.Value ?? 1)));
            payload.WriteByte(0);
            WriteColor(payload, font.Color);
            payload.WriteByte(ToFontScheme(font.GetFirstChild<FontScheme>()?.Val?.Value));
            WriteWideString(payload, name);
            XlsbRecordWriter.Write(output, BrtFont, payload.ToArray());
        }

        private static void WriteFill(Stream output, Fill fill) {
            GradientFill? gradient = fill.GetFirstChild<GradientFill>();
            if (gradient != null) {
                throw new NotSupportedException("Native XLSB generation does not yet encode gradient fill stops.");
            }

            PatternFill? pattern = fill.GetFirstChild<PatternFill>();
            using var payload = new MemoryStream(68);
            WriteUInt32(payload, ToFillPattern(pattern?.PatternType?.Value));
            WriteColor(payload, pattern?.ForegroundColor);
            WriteColor(payload, pattern?.BackgroundColor);
            WriteUInt32(payload, 0);
            for (int index = 0; index < 5; index++) WriteDouble(payload, 0D);
            WriteUInt32(payload, 0);
            XlsbRecordWriter.Write(output, BrtFill, payload.ToArray());
        }

        private static void WriteBorder(Stream output, Border border) {
            if (HasBorderContent(border.VerticalBorder)
                || HasBorderContent(border.HorizontalBorder)
                || HasBorderContent(border.StartBorder)
                || HasBorderContent(border.EndBorder)) {
                throw new NotSupportedException("Native XLSB generation does not yet encode start, end, vertical, or horizontal border definitions.");
            }

            using var payload = new MemoryStream(51);
            byte flags = 0;
            if (border.DiagonalDown?.Value == true) flags |= 0x01;
            if (border.DiagonalUp?.Value == true) flags |= 0x02;
            payload.WriteByte(flags);
            WriteBorderSide(payload, border.TopBorder);
            WriteBorderSide(payload, border.BottomBorder);
            WriteBorderSide(payload, border.LeftBorder);
            WriteBorderSide(payload, border.RightBorder);
            WriteBorderSide(payload, border.DiagonalBorder);
            XlsbRecordWriter.Write(output, BrtBorder, payload.ToArray());
        }

        private static bool HasBorderContent(BorderPropertiesType? border) =>
            border != null && (border.Style?.Value != null || border.Color != null);

        private static void WriteBorderSide(Stream output, BorderPropertiesType? side) {
            output.WriteByte(ToBorderStyle(side?.Style?.Value));
            output.WriteByte(0);
            WriteColor(output, side?.Color);
        }

        private static void WriteCellFormat(Stream output, CellFormat format, bool isStyleFormat) {
            using var payload = new MemoryStream(16);
            WriteUInt16(payload, isStyleFormat ? ushort.MaxValue : ToUInt16(format.FormatId?.Value ?? 0U, "parent style"));
            ushort numberFormatId = ToUInt16(format.NumberFormatId?.Value ?? 0U, "number format");
            ushort fontId = ToUInt16(format.FontId?.Value ?? 0U, "font");
            ushort fillId = ToUInt16(format.FillId?.Value ?? 0U, "fill");
            ushort borderId = ToUInt16(format.BorderId?.Value ?? 0U, "border");
            WriteUInt16(payload, numberFormatId);
            WriteUInt16(payload, fontId);
            WriteUInt16(payload, fillId);
            WriteUInt16(payload, borderId);

            Alignment? alignment = format.Alignment;
            payload.WriteByte(checked((byte)(alignment?.TextRotation?.Value ?? 0U)));
            payload.WriteByte(checked((byte)(alignment?.Indent?.Value ?? 0U)));
            byte alignmentFlags = (byte)(ToHorizontalAlignment(alignment?.Horizontal?.Value)
                | (ToVerticalAlignment(alignment?.Vertical?.Value) << 3));
            if (alignment?.WrapText?.Value == true) alignmentFlags |= 0x40;
            if (alignment?.JustifyLastLine?.Value == true) alignmentFlags |= 0x80;
            payload.WriteByte(alignmentFlags);

            byte protectionFlags = 0;
            if (alignment?.ShrinkToFit?.Value == true) protectionFlags |= 0x01;
            if (string.Equals(alignment?.MergeCell?.Value, "1", StringComparison.Ordinal)
                || string.Equals(alignment?.MergeCell?.Value, "true", StringComparison.OrdinalIgnoreCase)) protectionFlags |= 0x02;
            protectionFlags |= (byte)((alignment?.ReadingOrder?.Value ?? 0U) << 2);
            if (format.Protection?.Locked?.Value ?? true) protectionFlags |= 0x10;
            if (format.Protection?.Hidden?.Value == true) protectionFlags |= 0x20;
            if (format.PivotButton?.Value == true) protectionFlags |= 0x40;
            if (format.QuotePrefix?.Value == true) protectionFlags |= 0x80;
            payload.WriteByte(protectionFlags);

            ushort applyFlags = 0;
            if (ShouldApply(format.ApplyNumberFormat, numberFormatId != 0)) applyFlags |= 1 << 0;
            if (ShouldApply(format.ApplyFont, fontId != 0)) applyFlags |= 1 << 1;
            if (ShouldApply(format.ApplyAlignment, alignment != null)) applyFlags |= 1 << 2;
            if (ShouldApply(format.ApplyBorder, borderId != 0)) applyFlags |= 1 << 3;
            if (ShouldApply(format.ApplyFill, fillId != 0)) applyFlags |= 1 << 4;
            if (ShouldApply(format.ApplyProtection, format.Protection != null)) applyFlags |= 1 << 5;
            WriteUInt16(payload, applyFlags);
            XlsbRecordWriter.Write(output, BrtXf, payload.ToArray());
        }

        private static bool ShouldApply(BooleanValue? value, bool inferred) => value?.Value ?? inferred;

        private static void WriteNormalStyle(Stream output, CellStyles? styles, int styleFormatCount) {
            CellStyle? normal = styles?.Elements<CellStyle>().SingleOrDefault();
            uint formatId = normal?.FormatId?.Value ?? 0U;
            if (formatId >= styleFormatCount) {
                throw new NotSupportedException("Native XLSB generation found an out-of-range Normal style format reference.");
            }

            string name = normal?.Name?.Value ?? "Normal";
            ValidateStyleName(name, "cell style");
            using var payload = new MemoryStream(16 + name.Length * 2);
            WriteUInt32(payload, formatId);
            payload.WriteByte(1);
            payload.WriteByte(0);
            payload.WriteByte(0);
            payload.WriteByte(0);
            WriteWideString(payload, name);
            WriteUInt32Record(output, BrtBeginStyles, 1);
            XlsbRecordWriter.Write(output, BrtStyle, payload.ToArray());
            XlsbRecordWriter.Write(output, BrtEndStyles);
        }

        private static void WriteTableStyles(Stream output, TableStyles? styles) {
            string table = styles?.DefaultTableStyle?.Value ?? "TableStyleMedium2";
            string pivot = styles?.DefaultPivotStyle?.Value ?? "PivotStyleLight16";
            ValidateStyleName(table, "default table style");
            ValidateStyleName(pivot, "default pivot style");
            using var payload = new MemoryStream(12 + (table.Length + pivot.Length) * 2);
            WriteUInt32(payload, 0);
            WriteWideString(payload, table);
            WriteWideString(payload, pivot);
            XlsbRecordWriter.Write(output, BrtBeginTableStyles, payload.ToArray());
            XlsbRecordWriter.Write(output, BrtEndTableStyles);
        }

        private static void ValidateStyleName(string value, string description) {
            if (value.Length > MaximumStyleNameLength) {
                throw new NotSupportedException(
                    $"Native XLSB generation found a {description} name longer than {MaximumStyleNameLength} characters.");
            }
        }

        private static void WriteColor(Stream output, ColorType? color) {
            byte type = 0;
            byte index = 0;
            byte red = 0;
            byte green = 0;
            byte blue = 0;
            byte alpha = 0;
            if (color?.Rgb?.Value is string rgb && rgb.Length > 0) {
                string normalized = rgb.Length == 6 ? "FF" + rgb : rgb;
                if (normalized.Length != 8
                    || !uint.TryParse(normalized, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out uint argb)) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode style color '{rgb}'.");
                }
                type = 2;
                alpha = (byte)(argb >> 24);
                red = (byte)(argb >> 16);
                green = (byte)(argb >> 8);
                blue = (byte)argb;
            } else if (color?.Theme?.Value is uint theme) {
                type = 3;
                index = checked((byte)theme);
            } else if (color?.Indexed?.Value is uint indexed) {
                type = 1;
                index = checked((byte)indexed);
            } else if (color?.Auto?.Value == true) {
                type = 0;
                index = 0x40;
            }

            output.WriteByte((byte)((type << 1) | 0x01));
            output.WriteByte(index);
            WriteInt16(output, ToTint(color?.Tint?.Value ?? 0D));
            output.WriteByte(red);
            output.WriteByte(green);
            output.WriteByte(blue);
            output.WriteByte(alpha);
        }

        private static short ToTint(double value) {
            if (value < -1D || value > 1D) {
                throw new NotSupportedException($"Native XLSB generation cannot encode tint value {value}.");
            }
            return checked((short)Math.Round(value * (value < 0D ? 32768D : 32767D), MidpointRounding.AwayFromZero));
        }

        private static byte ToUnderline(UnderlineValues? value) {
            if (value == UnderlineValues.Single) return 1;
            if (value == UnderlineValues.Double) return 2;
            if (value == UnderlineValues.SingleAccounting) return 0x21;
            if (value == UnderlineValues.DoubleAccounting) return 0x22;
            return 0;
        }

        private static ushort ToScript(VerticalAlignmentRunValues? value) {
            if (value == VerticalAlignmentRunValues.Superscript) return 1;
            if (value == VerticalAlignmentRunValues.Subscript) return 2;
            return 0;
        }

        private static byte ToFontScheme(FontSchemeValues? value) {
            if (value == FontSchemeValues.Major) return 1;
            if (value == FontSchemeValues.Minor) return 2;
            return 0;
        }

        private static uint ToFillPattern(PatternValues? value) {
            if (value == PatternValues.Solid) return 1;
            if (value == PatternValues.MediumGray) return 2;
            if (value == PatternValues.DarkGray) return 3;
            if (value == PatternValues.LightGray) return 4;
            if (value == PatternValues.DarkHorizontal) return 5;
            if (value == PatternValues.DarkVertical) return 6;
            if (value == PatternValues.DarkDown) return 7;
            if (value == PatternValues.DarkUp) return 8;
            if (value == PatternValues.DarkGrid) return 9;
            if (value == PatternValues.DarkTrellis) return 10;
            if (value == PatternValues.LightHorizontal) return 11;
            if (value == PatternValues.LightVertical) return 12;
            if (value == PatternValues.LightDown) return 13;
            if (value == PatternValues.LightUp) return 14;
            if (value == PatternValues.LightGrid) return 15;
            if (value == PatternValues.LightTrellis) return 16;
            if (value == PatternValues.Gray125) return 17;
            if (value == PatternValues.Gray0625) return 18;
            return 0;
        }

        private static byte ToBorderStyle(BorderStyleValues? value) {
            if (value == BorderStyleValues.Thin) return 1;
            if (value == BorderStyleValues.Medium) return 2;
            if (value == BorderStyleValues.Dashed) return 3;
            if (value == BorderStyleValues.Dotted) return 4;
            if (value == BorderStyleValues.Thick) return 5;
            if (value == BorderStyleValues.Double) return 6;
            if (value == BorderStyleValues.Hair) return 7;
            if (value == BorderStyleValues.MediumDashed) return 8;
            if (value == BorderStyleValues.DashDot) return 9;
            if (value == BorderStyleValues.MediumDashDot) return 10;
            if (value == BorderStyleValues.DashDotDot) return 11;
            if (value == BorderStyleValues.MediumDashDotDot) return 12;
            if (value == BorderStyleValues.SlantDashDot) return 13;
            return 0;
        }

        private static byte ToHorizontalAlignment(HorizontalAlignmentValues? value) {
            if (value == HorizontalAlignmentValues.Left) return 1;
            if (value == HorizontalAlignmentValues.Center) return 2;
            if (value == HorizontalAlignmentValues.Right) return 3;
            if (value == HorizontalAlignmentValues.Fill) return 4;
            if (value == HorizontalAlignmentValues.Justify) return 5;
            if (value == HorizontalAlignmentValues.CenterContinuous) return 6;
            if (value == HorizontalAlignmentValues.Distributed) return 7;
            return 0;
        }

        private static byte ToVerticalAlignment(VerticalAlignmentValues? value) {
            if (value == VerticalAlignmentValues.Center) return 1;
            if (value == VerticalAlignmentValues.Bottom) return 2;
            if (value == VerticalAlignmentValues.Justify) return 3;
            if (value == VerticalAlignmentValues.Distributed) return 4;
            return 0;
        }

        private static ushort ToUInt16(uint value, string description) {
            if (value > ushort.MaxValue) {
                throw new NotSupportedException($"Native XLSB generation cannot encode {description} id {value}.");
            }
            return (ushort)value;
        }

        private static void WriteUInt32Record(Stream output, int recordType, uint value) {
            using var payload = new MemoryStream(4);
            WriteUInt32(payload, value);
            XlsbRecordWriter.Write(output, recordType, payload.ToArray());
        }

        private static void WriteWideString(Stream output, string value) {
            WriteUInt32(output, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void WriteDouble(Stream output, double value) {
            byte[] bytes = BitConverter.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void WriteInt16(Stream output, short value) => WriteUInt16(output, unchecked((ushort)value));

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
    }
}
