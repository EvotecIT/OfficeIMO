using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Writes the core BIFF12 formatting collections referenced by normal worksheet cells.</summary>
    internal static class XlsbStylesheetPartWriter {
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
        private const int MaximumStyleNameLength = 255;
        private const int MaximumFonts = 0xFFD3;
        private const int MaximumCellFormats = 0xFF96;
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

        /// <summary>Defines style zero even when a new workbook has no explicit formatting.</summary>
        internal static byte[] CreateForNewWorkbook(Stylesheet? stylesheet, out int cellFormatCount, long maximumBytes = int.MaxValue,
            IReadOnlyList<Font>? additionalFonts = null) =>
            Create(OfficeXmlNode.FromSdk(stylesheet ?? ExcelDocument.CreateDefaultStylesheet()), out cellFormatCount,
                maximumBytes, additionalFonts?.Select(font => OfficeXmlNode.FromSdk(font)));

        /// <summary>Encodes a generated SpreadsheetML stylesheet without loading the SDK model.</summary>
        internal static byte[] CreateFromMarkup(string stylesheetXml) {
            byte[] bytes = Encoding.UTF8.GetBytes(stylesheetXml);
            OfficeXmlTree tree = OfficeXmlTree.Parse(bytes, bytes.Length, new OfficeXmlNamespaces.Table());
            return Create(new OfficeXmlNode(tree, tree.Root), out _);
        }

        internal static byte[] Create(OfficeXmlNode stylesheet, out int cellFormatCount, long maximumBytes = int.MaxValue,
            IEnumerable<OfficeXmlNode>? additionalFonts = null) {
            if (stylesheet.IsNone) throw new ArgumentException("The XLSB stylesheet is missing.", nameof(stylesheet));
            ValidateUnsupportedCollections(stylesheet);

            List<OfficeXmlNode> numberFormats = Children(Child(stylesheet, "numFmts"), "numFmt");
            List<OfficeXmlNode> fonts = Children(Child(stylesheet, "fonts"), "font");
            if (additionalFonts != null) fonts.AddRange(additionalFonts);
            RequireItems(fonts, "fonts", MaximumFonts);
            List<OfficeXmlNode> fills = RequireItems(Children(Child(stylesheet, "fills"), "fill"), "fills");
            List<OfficeXmlNode> borders = RequireItems(Children(Child(stylesheet, "borders"), "border"), "borders");
            List<OfficeXmlNode> styleFormats = RequireItems(Children(Child(stylesheet, "cellStyleXfs"), "xf"), "cell style formats");
            List<OfficeXmlNode> cellFormats = RequireItems(Children(Child(stylesheet, "cellXfs"), "xf"), "cell formats", MaximumCellFormats);
            cellFormatCount = cellFormats.Count;

            ValidateCellFormatReferences(styleFormats, fonts.Count, fills.Count, borders.Count, isStyleFormat: true);
            ValidateCellFormatReferences(cellFormats, fonts.Count, fills.Count, borders.Count, isStyleFormat: false, styleFormats.Count);

            using var output = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(maximumBytes);
            XlsbRecordWriter.Write(output, BrtBeginStyleSheet);
            if (numberFormats.Count > 0) {
                WriteCollection(output, BrtBeginFmts, BrtEndFmts, numberFormats, WriteNumberFormat);
            }
            WriteCollection(output, BrtBeginFonts, BrtEndFonts, fonts, WriteFont);
            WriteCollection(output, BrtBeginFills, BrtEndFills, fills, WriteFill);
            WriteCollection(output, BrtBeginBorders, BrtEndBorders, borders, WriteBorder);
            WriteCollection(output, BrtBeginCellStyleXfs, BrtEndCellStyleXfs, styleFormats,
                (stream, format) => WriteCellFormat(stream, format, isStyleFormat: true));
            WriteCollection(output, BrtBeginCellXfs, BrtEndCellXfs, cellFormats,
                (stream, format) => WriteCellFormat(stream, format, isStyleFormat: false));
            WriteCellStyles(output, Child(stylesheet, "cellStyles"), styleFormats.Count);
            List<OfficeXmlNode> differentialFormats = Children(Child(stylesheet, "dxfs"), "dxf");
            WriteCollection(output, BrtBeginDxfs, BrtEndDxfs, differentialFormats, (stream, format) => {
                if (format.SdkElement is not DifferentialFormat differentialFormat)
                    throw new NotSupportedException("Native XLSB generation needs the differential-format encoder for this stylesheet.");
                XlsbGeneratedRecord record = Styles.XlsbDifferentialFormatCodec.CreateRecord(differentialFormat);
                XlsbRecordWriter.Write(stream, record.Type, record.Payload);
            });
            WriteTableStyles(output, Child(stylesheet, "tableStyles"));
            XlsbRecordWriter.Write(output, BrtEndStyleSheet);
            return output.ToArray();
        }

        private static OfficeXmlNode Child(OfficeXmlNode parent, string name) => parent.Child(Spreadsheet, name);

        private static List<OfficeXmlNode> Children(OfficeXmlNode parent, string name) => parent.Children(Spreadsheet, name).ToList();

        private static void ValidateUnsupportedCollections(OfficeXmlNode stylesheet) {
            OfficeXmlNode differentialFormats = Child(stylesheet, "dxfs");
            if (!differentialFormats.IsNone
                && (differentialFormats.Children().Count != differentialFormats.Children(Spreadsheet, "dxf").Count
                    || XlsbWorkbookMarkup.UnsupportedAttribute(differentialFormats, new[] { "count" }) != null)) {
                throw new NotSupportedException("Native XLSB generation does not encode extensions on the differential formatting collection.");
            }
            bool hasStyleExtension = false;
            foreach (OfficeXmlNode child in stylesheet.Children()) {
                if (child.LocalName == "extLst") {
                    hasStyleExtension = true;
                    break;
                }
            }
            if (Child(stylesheet, "colors").Exists || hasStyleExtension) {
                throw new NotSupportedException("Native XLSB generation does not yet encode custom indexed colors or style extensions.");
            }

            if (Child(stylesheet, "tableStyles").FirstChild.Exists) {
                throw new NotSupportedException("Native XLSB generation does not yet encode custom table or pivot styles.");
            }
        }

        private static List<OfficeXmlNode> RequireItems(List<OfficeXmlNode> values, string collectionName, int maximumItems = 65_536) {
            if (values.Count == 0) {
                throw new NotSupportedException($"Native XLSB generation requires a non-empty {collectionName} collection.");
            }
            if (values.Count > maximumItems) {
                throw new NotSupportedException($"Native XLSB generation supports at most {maximumItems:N0} {collectionName}.");
            }
            return values;
        }

        private static void ValidateCellFormatReferences(
            IEnumerable<OfficeXmlNode> formats,
            int fontCount,
            int fillCount,
            int borderCount,
            bool isStyleFormat,
            int styleFormatCount = 0) {
            foreach (OfficeXmlNode format in formats) {
                uint fontId = XlsbWorkbookMarkup.ReadUInt32(format, "fontId") ?? 0U;
                uint fillId = XlsbWorkbookMarkup.ReadUInt32(format, "fillId") ?? 0U;
                uint borderId = XlsbWorkbookMarkup.ReadUInt32(format, "borderId") ?? 0U;
                uint parentId = XlsbWorkbookMarkup.ReadUInt32(format, "xfId") ?? 0U;
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

        internal static XlsbGeneratedRecord CreateNumberFormatRecord(NumberingFormat value) {
            using var output = new MemoryStream();
            WriteNumberFormat(output, OfficeXmlNode.FromSdk(value));
            output.Position = 0;
            var record = XlsbRecordReader.ReadAll(output).Single();
            return new XlsbGeneratedRecord(record.Type, record.Data);
        }

        private static void WriteNumberFormat(Stream output, OfficeXmlNode format) {
            uint id = XlsbWorkbookMarkup.ReadUInt32(format, "numFmtId")
                ?? throw new NotSupportedException("Native XLSB generation requires every custom number format to have an id.");
            string code = format.Attribute("formatCode") ?? string.Empty;
            if (id > ushort.MaxValue || string.IsNullOrEmpty(code) || code.Length > 255) {
                throw new NotSupportedException("Native XLSB generation found an invalid custom number format.");
            }

            using var payload = new MemoryStream(8 + code.Length * 2);
            WriteUInt16(payload, (ushort)id);
            WriteWideString(payload, code);
            XlsbRecordWriter.Write(output, BrtFmt, payload.ToArray());
        }

        private static void WriteFont(Stream output, OfficeXmlNode font) {
            XlsbGeneratedRecord record = CreateFontRecord(font);
            XlsbRecordWriter.Write(output, record.Type, record.Payload);
        }

        internal static XlsbGeneratedRecord CreateFontRecord(Font font) => CreateFontRecord(OfficeXmlNode.FromSdk(font));

        private static XlsbGeneratedRecord CreateFontRecord(OfficeXmlNode font) {
            string name = Child(font, "name").Attribute("val") ?? "Calibri";
            double size = ReadDouble(Child(font, "sz"), "val") ?? 11D;
            if (string.IsNullOrWhiteSpace(name) || name.Length > 31 || size <= 0D || size > 409.55D) {
                throw new NotSupportedException("Native XLSB generation found a font with an unsupported name or size.");
            }

            bool bold = FontFlag(font, "b");
            ushort flags = 0;
            if (bold) flags |= 0x0001;
            if (FontFlag(font, "i")) flags |= 0x0002;
            if (FontFlag(font, "strike")) flags |= 0x0008;
            if (FontFlag(font, "outline")) flags |= 0x0010;
            if (FontFlag(font, "shadow")) flags |= 0x0020;
            if (FontFlag(font, "condense")) flags |= 0x0040;
            if (FontFlag(font, "extend")) flags |= 0x0080;

            using var payload = new MemoryStream(32 + name.Length * 2);
            WriteUInt16(payload, checked((ushort)Math.Round(size * 20D, MidpointRounding.AwayFromZero)));
            WriteUInt16(payload, flags);
            WriteUInt16(payload, bold ? (ushort)700 : (ushort)400);
            WriteUInt16(payload, Child(font, "vertAlign").Attribute("val") switch { "superscript" => (ushort)1, "subscript" => (ushort)2, _ => (ushort)0 });
            OfficeXmlNode underline = Child(font, "u");
            payload.WriteByte(underline.IsNone ? (byte)0 : (underline.Attribute("val") ?? "single") switch {
                "single" => 1, "double" => 2, "singleAccounting" => 0x21, "doubleAccounting" => 0x22, _ => 0
            });
            payload.WriteByte(checked((byte)(XlsbWorkbookMarkup.ReadInt32(Child(font, "family"), "val") ?? 0)));
            payload.WriteByte(checked((byte)(XlsbWorkbookMarkup.ReadInt32(Child(font, "charset"), "val") ?? 1)));
            payload.WriteByte(0);
            WriteColor(payload, Child(font, "color"));
            payload.WriteByte(Child(font, "scheme").Attribute("val") switch { "major" => 1, "minor" => 2, _ => 0 });
            WriteWideString(payload, name);
            return new XlsbGeneratedRecord(BrtFont, payload.ToArray());
        }

        internal static XlsbGeneratedRecord CreateFillRecord(Fill value) {
            using var output = new MemoryStream();
            WriteFill(output, OfficeXmlNode.FromSdk(value));
            output.Position = 0;
            var record = XlsbRecordReader.ReadAll(output).Single();
            return new XlsbGeneratedRecord(record.Type, record.Data);
        }

        private static void WriteFill(Stream output, OfficeXmlNode fill) {
            if (Child(fill, "gradientFill").Exists) {
                throw new NotSupportedException("Native XLSB generation does not yet encode gradient fill stops.");
            }

            OfficeXmlNode pattern = Child(fill, "patternFill");
            using var payload = new MemoryStream(68);
            WriteUInt32(payload, XlsbStylesheetMarkup.FillPatternCode(pattern.Attribute("patternType")));
            WriteColor(payload, Child(pattern, "fgColor"));
            WriteColor(payload, Child(pattern, "bgColor"));
            WriteUInt32(payload, 0);
            for (int index = 0; index < 5; index++) WriteDouble(payload, 0D);
            WriteUInt32(payload, 0);
            XlsbRecordWriter.Write(output, BrtFill, payload.ToArray());
        }

        internal static XlsbGeneratedRecord CreateBorderRecord(Border value) {
            using var output = new MemoryStream();
            WriteBorder(output, OfficeXmlNode.FromSdk(value));
            output.Position = 0;
            var record = XlsbRecordReader.ReadAll(output).Single();
            return new XlsbGeneratedRecord(record.Type, record.Data);
        }

        private static void WriteBorder(Stream output, OfficeXmlNode border) {
            if (HasBorderContent(Child(border, "vertical"))
                || HasBorderContent(Child(border, "horizontal"))
                || HasBorderContent(Child(border, "start"))
                || HasBorderContent(Child(border, "end"))) {
                throw new NotSupportedException("Native XLSB generation does not yet encode start, end, vertical, or horizontal border definitions.");
            }

            using var payload = new MemoryStream(51);
            byte flags = 0;
            if (XlsbWorkbookMarkup.ReadBoolean(border, "diagonalDown") == true) flags |= 0x01;
            if (XlsbWorkbookMarkup.ReadBoolean(border, "diagonalUp") == true) flags |= 0x02;
            payload.WriteByte(flags);
            WriteBorderSide(payload, Child(border, "top"));
            WriteBorderSide(payload, Child(border, "bottom"));
            WriteBorderSide(payload, Child(border, "left"));
            WriteBorderSide(payload, Child(border, "right"));
            WriteBorderSide(payload, Child(border, "diagonal"));
            XlsbRecordWriter.Write(output, BrtBorder, payload.ToArray());
        }

        private static bool HasBorderContent(OfficeXmlNode border) =>
            border.HasAttribute("style") || Child(border, "color").Exists;

        private static void WriteBorderSide(Stream output, OfficeXmlNode side) {
            output.WriteByte(XlsbStylesheetMarkup.BorderStyleCode(side.Attribute("style")));
            output.WriteByte(0);
            WriteColor(output, Child(side, "color"));
        }

        private static void WriteCellFormat(Stream output, OfficeXmlNode format, bool isStyleFormat) {
            XlsbGeneratedRecord record = CreateCellFormatRecord(format, isStyleFormat);
            XlsbRecordWriter.Write(output, record.Type, record.Payload);
        }

        internal static XlsbGeneratedRecord CreateCellFormatRecord(CellFormat format, bool isStyleFormat = false) =>
            CreateCellFormatRecord(OfficeXmlNode.FromSdk(format), isStyleFormat);

        private static XlsbGeneratedRecord CreateCellFormatRecord(OfficeXmlNode format, bool isStyleFormat) {
            using var payload = new MemoryStream(16);
            WriteUInt16(payload, isStyleFormat ? ushort.MaxValue : ToUInt16(XlsbWorkbookMarkup.ReadUInt32(format, "xfId") ?? 0U, "parent style"));
            ushort numberFormatId = ToUInt16(XlsbWorkbookMarkup.ReadUInt32(format, "numFmtId") ?? 0U, "number format");
            ushort fontId = ToUInt16(XlsbWorkbookMarkup.ReadUInt32(format, "fontId") ?? 0U, "font");
            ushort fillId = ToUInt16(XlsbWorkbookMarkup.ReadUInt32(format, "fillId") ?? 0U, "fill");
            ushort borderId = ToUInt16(XlsbWorkbookMarkup.ReadUInt32(format, "borderId") ?? 0U, "border");
            WriteUInt16(payload, numberFormatId);
            WriteUInt16(payload, fontId);
            WriteUInt16(payload, fillId);
            WriteUInt16(payload, borderId);

            OfficeXmlNode alignment = Child(format, "alignment");
            OfficeXmlNode protection = Child(format, "protection");
            payload.WriteByte(checked((byte)(XlsbWorkbookMarkup.ReadUInt32(alignment, "textRotation") ?? 0U)));
            payload.WriteByte(checked((byte)(XlsbWorkbookMarkup.ReadUInt32(alignment, "indent") ?? 0U)));
            byte alignmentFlags = (byte)(XlsbStylesheetMarkup.HorizontalAlignmentCode(alignment.Attribute("horizontal"))
                | (XlsbStylesheetMarkup.VerticalAlignmentCode(alignment.Attribute("vertical")) << 3));
            if (XlsbWorkbookMarkup.ReadBoolean(alignment, "wrapText") == true) alignmentFlags |= 0x40;
            if (XlsbWorkbookMarkup.ReadBoolean(alignment, "justifyLastLine") == true) alignmentFlags |= 0x80;
            payload.WriteByte(alignmentFlags);

            byte protectionFlags = 0;
            if (XlsbWorkbookMarkup.ReadBoolean(alignment, "shrinkToFit") == true) protectionFlags |= 0x01;
            if (string.Equals(alignment.Attribute("mergeCell"), "1", StringComparison.Ordinal)
                || string.Equals(alignment.Attribute("mergeCell"), "true", StringComparison.OrdinalIgnoreCase)) protectionFlags |= 0x02;
            protectionFlags |= (byte)((XlsbWorkbookMarkup.ReadUInt32(alignment, "readingOrder") ?? 0U) << 2);
            if (XlsbWorkbookMarkup.ReadBoolean(protection, "locked") ?? true) protectionFlags |= 0x10;
            if (XlsbWorkbookMarkup.ReadBoolean(protection, "hidden") == true) protectionFlags |= 0x20;
            if (XlsbWorkbookMarkup.ReadBoolean(format, "pivotButton") == true) protectionFlags |= 0x40;
            if (XlsbWorkbookMarkup.ReadBoolean(format, "quotePrefix") == true) protectionFlags |= 0x80;
            payload.WriteByte(protectionFlags);

            ushort applyFlags = 0;
            if (XlsbWorkbookMarkup.ReadBoolean(format, "applyNumberFormat") ?? numberFormatId != 0) applyFlags |= 1 << 0;
            if (XlsbWorkbookMarkup.ReadBoolean(format, "applyFont") ?? fontId != 0) applyFlags |= 1 << 1;
            if (XlsbWorkbookMarkup.ReadBoolean(format, "applyAlignment") ?? alignment.Exists) applyFlags |= 1 << 2;
            if (XlsbWorkbookMarkup.ReadBoolean(format, "applyBorder") ?? borderId != 0) applyFlags |= 1 << 3;
            if (XlsbWorkbookMarkup.ReadBoolean(format, "applyFill") ?? fillId != 0) applyFlags |= 1 << 4;
            if (XlsbWorkbookMarkup.ReadBoolean(format, "applyProtection") ?? protection.Exists) applyFlags |= 1 << 5;
            WriteUInt16(payload, applyFlags);
            return new XlsbGeneratedRecord(BrtXf, payload.ToArray());
        }

        private static double? ReadDouble(OfficeXmlNode node, string attribute) =>
            node.Attribute(attribute) is string value ? XmlConvert.ToDouble(value) : null;

        private static bool FontFlag(OfficeXmlNode font, string localName) {
            OfficeXmlNode child = Child(font, localName);
            return child.Exists && (XlsbWorkbookMarkup.ReadBoolean(child, "val") ?? true);
        }

        // BrtStyle (MS-XLSB 2.4.811): ixf, StyleFlags (fBuiltIn, fHidden, fCustom), iStyBuiltIn, iLevel, stName.
        private static void WriteCellStyles(Stream output, OfficeXmlNode styles, int styleFormatCount) {
            List<OfficeXmlNode> items = Children(styles, "cellStyle");
            if (items.Count == 0) items.Add(default); // Built-in Normal when no cell styles were declared.
            var formatIds = new HashSet<uint>();
            var payloads = new List<byte[]>(items.Count);
            foreach (OfficeXmlNode style in items) {
                if (style.FirstChild.Exists) {
                    throw new NotSupportedException($"Native XLSB generation does not yet encode extensions on cell style '{style.Attribute("name")}'.");
                }
                uint formatId = XlsbWorkbookMarkup.ReadUInt32(style, "xfId") ?? 0U;
                if (formatId >= styleFormatCount) {
                    throw new NotSupportedException("Native XLSB generation found an out-of-range cell style format reference.");
                }
                if (!formatIds.Add(formatId)) {
                    throw new NotSupportedException($"Native XLSB generation requires each cell style to use its own style format; format {formatId} is shared.");
                }
                uint? builtInId = style.IsNone ? 0U : XlsbWorkbookMarkup.ReadUInt32(style, "builtinId");
                if (builtInId > byte.MaxValue) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode built-in cell style id {builtInId}.");
                }
                bool customBuiltIn = XlsbWorkbookMarkup.ReadBoolean(style, "customBuiltin") == true;
                if (customBuiltIn && !builtInId.HasValue) {
                    throw new NotSupportedException("Native XLSB generation requires a built-in id for a customized built-in cell style.");
                }
                uint level = XlsbWorkbookMarkup.ReadUInt32(style, "iLevel") ?? 0U;
                if (builtInId is 1U or 2U && level > 6U) {
                    throw new NotSupportedException($"Native XLSB generation supports outline cell-style levels from 0 through 6, not {level}.");
                }
                string name = style.Attribute("name") ?? (builtInId == 0U ? "Normal" : string.Empty);
                if (name.Length == 0) throw new NotSupportedException("Native XLSB generation requires a name for every cell style.");
                ValidateStyleName(name, "cell style");

                ushort flags = (ushort)((builtInId.HasValue ? 0x0001 : 0)
                    | (XlsbWorkbookMarkup.ReadBoolean(style, "hidden") == true ? 0x0002 : 0)
                    | (customBuiltIn ? 0x0004 : 0));
                using var payload = new MemoryStream(16 + name.Length * 2);
                WriteUInt32(payload, formatId);
                WriteUInt16(payload, flags);
                payload.WriteByte((byte)(builtInId ?? 0U));
                payload.WriteByte(builtInId is 1U or 2U ? (byte)level : (byte)0);
                WriteWideString(payload, name);
                payloads.Add(payload.ToArray());
            }

            WriteUInt32Record(output, BrtBeginStyles, checked((uint)payloads.Count));
            foreach (byte[] payload in payloads) XlsbRecordWriter.Write(output, BrtStyle, payload);
            XlsbRecordWriter.Write(output, BrtEndStyles);
        }

        private static void WriteTableStyles(Stream output, OfficeXmlNode styles) {
            string table = styles.Attribute("defaultTableStyle") ?? "TableStyleMedium2";
            string pivot = styles.Attribute("defaultPivotStyle") ?? "PivotStyleLight16";
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

        /// <summary>
        /// Writes a BrtColor-layout color (MS-XLSB 2.4.337), which XFPropColor (2.5.161) shares; XFPropColor forbids the
        /// full-shade tint -32768, so <paramref name="allowFullShade"/> false writes -32767 for a tint of -1.
        /// </summary>
        internal static void WriteColor(Stream output, ColorType? color, bool allowFullShade = true) =>
            WriteColor(output, OfficeXmlNode.FromSdk(color), allowFullShade);

        private static void WriteColor(Stream output, OfficeXmlNode color, bool allowFullShade = true) {
            byte type = 0;
            byte index = 0;
            byte red = 0;
            byte green = 0;
            byte blue = 0;
            byte alpha = 0;
            if (color.Attribute("rgb") is string rgb && rgb.Length > 0) {
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
            } else if (XlsbWorkbookMarkup.ReadUInt32(color, "theme") is uint theme) {
                type = 3;
                index = checked((byte)theme);
            } else if (XlsbWorkbookMarkup.ReadUInt32(color, "indexed") is uint indexed) {
                type = 1;
                index = checked((byte)indexed);
            } else if (XlsbWorkbookMarkup.ReadBoolean(color, "auto") == true) {
                type = 0;
                index = 0x40;
            }

            output.WriteByte((byte)((type << 1) | 0x01));
            output.WriteByte(index);
            short tint = ToTint(ReadDouble(color, "tint") ?? 0D);
            if (!allowFullShade && tint == short.MinValue) tint = -32767;
            WriteInt16(output, tint);
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

        internal static uint ToFillPattern(PatternValues? value) {
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

        internal static byte ToBorderStyle(BorderStyleValues? value) {
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

        internal static byte ToHorizontalAlignment(HorizontalAlignmentValues? value) {
            if (value == HorizontalAlignmentValues.Left) return 1;
            if (value == HorizontalAlignmentValues.Center) return 2;
            if (value == HorizontalAlignmentValues.Right) return 3;
            if (value == HorizontalAlignmentValues.Fill) return 4;
            if (value == HorizontalAlignmentValues.Justify) return 5;
            if (value == HorizontalAlignmentValues.CenterContinuous) return 6;
            if (value == HorizontalAlignmentValues.Distributed) return 7;
            return 0;
        }

        // MS-XLSB BrtXF alcV: 0 is top and 2 is bottom, the SpreadsheetML default for a missing value.
        internal static byte ToVerticalAlignment(VerticalAlignmentValues? value) {
            if (value == VerticalAlignmentValues.Top) return 0;
            if (value == VerticalAlignmentValues.Center) return 1;
            if (value == VerticalAlignmentValues.Justify) return 3;
            if (value == VerticalAlignmentValues.Distributed) return 4;
            return 2;
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
