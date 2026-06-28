using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal sealed class LegacyXlsStyleTable {
        private readonly Dictionary<uint, ushort> _styleIndexMap;

        private LegacyXlsStyleTable(
            Dictionary<uint, ushort> styleIndexMap,
            IReadOnlyList<byte[]> formatRecords,
            IReadOnlyList<byte[]> cellFormatRecords) {
            _styleIndexMap = styleIndexMap;
            FormatRecords = formatRecords;
            CellFormatRecords = cellFormatRecords;
        }

        internal IReadOnlyList<byte[]> FormatRecords { get; }

        internal IReadOnlyList<byte[]> CellFormatRecords { get; }

        internal static LegacyXlsStyleTable Create(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets, LegacyXlsFontTable fontTable) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));
            if (fontTable == null) throw new ArgumentNullException(nameof(fontTable));

            SortedSet<uint> openXmlStyleIndexes = CollectReferencedStyleIndexes(sheets);
            openXmlStyleIndexes.Add(0U);

            Stylesheet? stylesheet = document.WorkbookPartRoot?.WorkbookStylesPart?.Stylesheet;
            List<CellFormat> openXmlCellFormats = stylesheet?.CellFormats?.Elements<CellFormat>().ToList() ?? new List<CellFormat>();
            List<CellFormat> openXmlCellStyleFormats = stylesheet?.CellStyleFormats?.Elements<CellFormat>().ToList() ?? new List<CellFormat>();
            List<Font> openXmlFonts = stylesheet?.Fonts?.Elements<Font>().ToList() ?? new List<Font>();
            List<Fill> openXmlFills = stylesheet?.Fills?.Elements<Fill>().ToList() ?? new List<Fill>();
            List<Border> openXmlBorders = stylesheet?.Borders?.Elements<Border>().ToList() ?? new List<Border>();
            Dictionary<uint, string> customNumberFormats = stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
                .Where(format => format.NumberFormatId?.Value is uint && format.FormatCode?.Value != null)
                .ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode!.Value!)
                ?? new Dictionary<uint, string>();

            var map = new Dictionary<uint, ushort>();
            var formatRecordsById = new Dictionary<ushort, byte[]>();
            var cellFormatRecords = new List<byte[]>();

            foreach (uint openXmlStyleIndex in openXmlStyleIndexes) {
                CellFormat? cellFormat = openXmlStyleIndex < openXmlCellFormats.Count
                    ? openXmlCellFormats[(int)openXmlStyleIndex]
                    : null;

                if (openXmlStyleIndex != 0U && cellFormat == null) {
                    throw new NotSupportedException($"Native XLS saving cannot resolve Open XML style index {openXmlStyleIndex}.");
                }

                CellFormat? parentCellFormat = ResolveParentCellFormat(cellFormat, openXmlCellStyleFormats);
                ushort numberFormatId = ResolveNumberFormatId(openXmlStyleIndex, cellFormat, parentCellFormat, customNumberFormats, formatRecordsById);
                ushort fontIndex = ResolveFontIndex(openXmlStyleIndex, cellFormat, parentCellFormat, openXmlFonts, fontTable);
                bool applyFont = fontIndex != 0 || cellFormat?.ApplyFont?.Value == true;
                FillInfo fill = ResolveFill(openXmlStyleIndex, cellFormat, parentCellFormat, openXmlFills, fontTable);
                AlignmentInfo alignment = ResolveAlignment(openXmlStyleIndex, cellFormat, parentCellFormat);
                BorderInfo border = ResolveBorder(openXmlStyleIndex, cellFormat, parentCellFormat, openXmlBorders, fontTable);
                ProtectionInfo protection = ResolveProtection(cellFormat, parentCellFormat);
                map[openXmlStyleIndex] = checked((ushort)cellFormatRecords.Count);
                cellFormatRecords.Add(BuildXfPayload(numberFormatId, fontIndex, applyFont, fill, alignment, border, protection));
            }

            return new LegacyXlsStyleTable(
                map,
                formatRecordsById.OrderBy(pair => pair.Key).Select(pair => pair.Value).ToArray(),
                cellFormatRecords);
        }

        internal ushort GetBiffStyleIndex(uint? openXmlStyleIndex) {
            uint index = openXmlStyleIndex ?? 0U;
            return _styleIndexMap.TryGetValue(index, out ushort biffStyleIndex)
                ? biffStyleIndex
                : throw new NotSupportedException($"Native XLS saving cannot resolve Open XML style index {index}.");
        }

        private static SortedSet<uint> CollectReferencedStyleIndexes(IReadOnlyList<ExcelSheet> sheets) {
            var indexes = new SortedSet<uint>();
            foreach (ExcelSheet sheet in sheets) {
                foreach (ExcelColumnSnapshot column in sheet.GetColumnDefinitions()) {
                    if (column.StyleIndex.HasValue) {
                        indexes.Add(column.StyleIndex.Value);
                    }
                }

                foreach (ExcelRowSnapshot row in sheet.GetRowDefinitions()) {
                    if (row.StyleIndex.HasValue) {
                        indexes.Add(row.StyleIndex.Value);
                    }
                }

                SheetData? sheetData = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetData>();
                if (sheetData == null) {
                    continue;
                }

                foreach (Row row in sheetData.Elements<Row>()) {
                    foreach (Cell cell in row.Elements<Cell>()) {
                        if (cell.StyleIndex?.Value is uint styleIndex) {
                            indexes.Add(styleIndex);
                        }
                    }
                }
            }

            return indexes;
        }

        private static CellFormat? ResolveParentCellFormat(CellFormat? cellFormat, IReadOnlyList<CellFormat> openXmlCellStyleFormats) {
            if (cellFormat?.FormatId?.Value is not uint parentIndex || parentIndex >= openXmlCellStyleFormats.Count) {
                return null;
            }

            return openXmlCellStyleFormats[(int)parentIndex];
        }

        private static ushort ResolveNumberFormatId(
            uint openXmlStyleIndex,
            CellFormat? cellFormat,
            CellFormat? parentCellFormat,
            IReadOnlyDictionary<uint, string> customNumberFormats,
            Dictionary<ushort, byte[]> formatRecordsById) {
            if (cellFormat == null && parentCellFormat == null) {
                return 0;
            }

            ThrowIfUnsupportedStyleFacet(openXmlStyleIndex, cellFormat);
            ThrowIfUnsupportedStyleFacet(openXmlStyleIndex, parentCellFormat);

            uint numberFormatId = cellFormat?.NumberFormatId?.Value ?? parentCellFormat?.NumberFormatId?.Value ?? 0U;
            if (numberFormatId > ushort.MaxValue) {
                throw new NotSupportedException($"Native XLS saving supports BIFF8 number format ids up to {ushort.MaxValue}.");
            }

            ushort biffNumberFormatId = checked((ushort)numberFormatId);
            if (biffNumberFormatId == 0 || BiffBuiltInNumberFormat.TryGetCode(biffNumberFormatId, out _)) {
                return biffNumberFormatId;
            }

            if (!customNumberFormats.TryGetValue(numberFormatId, out string? formatCode) || string.IsNullOrWhiteSpace(formatCode)) {
                throw new NotSupportedException($"Native XLS saving cannot resolve custom number format id {numberFormatId} for Open XML style index {openXmlStyleIndex}.");
            }

            if (!formatRecordsById.TryGetValue(biffNumberFormatId, out byte[]? existingRecord)) {
                formatRecordsById.Add(biffNumberFormatId, BuildFormatPayload(biffNumberFormatId, formatCode));
            } else if (!FormatRecordMatches(existingRecord, biffNumberFormatId, formatCode)) {
                throw new NotSupportedException($"Native XLS saving encountered conflicting custom number format definitions for id {numberFormatId}.");
            }

            return biffNumberFormatId;
        }

        private static ushort ResolveFontIndex(
            uint openXmlStyleIndex,
            CellFormat? cellFormat,
            CellFormat? parentCellFormat,
            IReadOnlyList<Font> openXmlFonts,
            LegacyXlsFontTable fontTable) {
            UInt32Value? fontIdValue = cellFormat?.FontId ?? parentCellFormat?.FontId;
            if (fontIdValue == null) {
                return 0;
            }

            uint fontId = fontIdValue.Value;
            if (fontId >= openXmlFonts.Count) {
                if (fontId == 0U) {
                    return 0;
                }

                throw new NotSupportedException($"Native XLS saving cannot resolve font id {fontId} for Open XML style index {openXmlStyleIndex}.");
            }

            Font font = openXmlFonts[(int)fontId];
            ThrowIfStyleExtensionsPresent(openXmlStyleIndex, font, "font style");
            if (!fontTable.TryGetFontIndex(font, out ushort fontIndex, out string? reason)) {
                throw new NotSupportedException($"Native XLS saving does not yet support {reason ?? "cell font settings"} for Open XML style index {openXmlStyleIndex}.");
            }

            return fontIndex;
        }

        private static FillInfo ResolveFill(
            uint openXmlStyleIndex,
            CellFormat? cellFormat,
            CellFormat? parentCellFormat,
            IReadOnlyList<Fill> openXmlFills,
            LegacyXlsFontTable fontTable) {
            UInt32Value? fillIdValue = cellFormat?.FillId ?? parentCellFormat?.FillId;
            if (fillIdValue == null || fillIdValue.Value == 0U) {
                return FillInfo.None;
            }

            uint fillId = fillIdValue.Value;
            if (fillId >= openXmlFills.Count) {
                throw new NotSupportedException($"Native XLS saving cannot resolve fill id {fillId} for Open XML style index {openXmlStyleIndex}.");
            }

            Fill fill = openXmlFills[(int)fillId];
            ThrowIfStyleExtensionsPresent(openXmlStyleIndex, fill, "fill style");
            ThrowIfMultipleStyleChildren<GradientFill>(openXmlStyleIndex, fill, "fill style", "gradient fill");
            ThrowIfMultipleStyleChildren<PatternFill>(openXmlStyleIndex, fill, "fill style", "pattern fill");
            if (fill.GradientFill != null) {
                throw new NotSupportedException($"Native XLS saving does not yet support gradient fills for Open XML style index {openXmlStyleIndex}.");
            }

            PatternFill? patternFill = fill.PatternFill;
            PatternValues? patternType = patternFill?.PatternType?.Value;
            if (patternFill == null || patternType == null || patternType == PatternValues.None) {
                return FillInfo.None;
            }

            ThrowIfMultipleStyleChildren<ForegroundColor>(openXmlStyleIndex, patternFill, "fill style", "foreground color");
            ThrowIfMultipleStyleChildren<BackgroundColor>(openXmlStyleIndex, patternFill, "fill style", "background color");

            byte pattern = ToFillPatternCode(openXmlStyleIndex, patternType.Value);

            ColorType? foregroundColor = patternFill.ForegroundColor ?? (ColorType?)patternFill.BackgroundColor;
            ColorType? backgroundColor = patternFill.BackgroundColor ?? foregroundColor;
            if (foregroundColor == null || backgroundColor == null) {
                return FillInfo.None;
            }

            ushort foregroundIndex = ResolveFillColorIndex(openXmlStyleIndex, foregroundColor, "foreground", fontTable);
            ushort backgroundIndex = ResolveFillColorIndex(openXmlStyleIndex, backgroundColor, "background", fontTable);
            return new FillInfo(pattern, foregroundIndex, backgroundIndex);
        }

        private static byte ToFillPatternCode(uint openXmlStyleIndex, PatternValues pattern) {
            if (pattern == PatternValues.Solid) return 1;
            if (pattern == PatternValues.MediumGray) return 2;
            if (pattern == PatternValues.DarkGray) return 3;
            if (pattern == PatternValues.LightGray) return 4;
            if (pattern == PatternValues.DarkHorizontal) return 5;
            if (pattern == PatternValues.DarkVertical) return 6;
            if (pattern == PatternValues.DarkDown) return 7;
            if (pattern == PatternValues.DarkUp) return 8;
            if (pattern == PatternValues.DarkGrid) return 9;
            if (pattern == PatternValues.DarkTrellis) return 10;
            if (pattern == PatternValues.LightHorizontal) return 11;
            if (pattern == PatternValues.LightVertical) return 12;
            if (pattern == PatternValues.LightDown) return 13;
            if (pattern == PatternValues.LightUp) return 14;
            if (pattern == PatternValues.LightGrid) return 15;
            if (pattern == PatternValues.LightTrellis) return 16;
            if (pattern == PatternValues.Gray125) return 17;
            if (pattern == PatternValues.Gray0625) return 18;
            throw new NotSupportedException($"Native XLS saving does not yet support fill pattern '{pattern}' for Open XML style index {openXmlStyleIndex}.");
        }

        private static ushort ResolveFillColorIndex(
            uint openXmlStyleIndex,
            ColorType color,
            string role,
            LegacyXlsFontTable fontTable) {
            if (!fontTable.TryGetColorIndex(color, $"cell fill {role}", "colors", out ushort colorIndex, out string? reason)) {
                throw new NotSupportedException($"Native XLS saving does not yet support {reason ?? "cell fill colors"} for Open XML style index {openXmlStyleIndex}.");
            }

            if (colorIndex > 0x007f) {
                throw new NotSupportedException($"Native XLS saving does not yet support cell fill {role} color indexes above 0x007F for Open XML style index {openXmlStyleIndex}.");
            }

            return colorIndex;
        }

        private static AlignmentInfo ResolveAlignment(uint openXmlStyleIndex, CellFormat? cellFormat, CellFormat? parentCellFormat) {
            Alignment? alignment = cellFormat?.Alignment ?? parentCellFormat?.Alignment;
            if (alignment == null) {
                return AlignmentInfo.None;
            }

            ThrowIfStyleExtensionsPresent(openXmlStyleIndex, alignment, "alignment style");
            byte horizontal = ToHorizontalAlignmentCode(openXmlStyleIndex, alignment.Horizontal?.Value);
            byte vertical = ToVerticalAlignmentCode(openXmlStyleIndex, alignment.Vertical?.Value);
            bool wrapText = alignment.WrapText?.Value == true;
            byte textRotation = ToTextRotationCode(openXmlStyleIndex, alignment.TextRotation?.Value);
            byte indent = ToIndentCode(openXmlStyleIndex, alignment.Indent?.Value);
            bool shrinkToFit = alignment.ShrinkToFit?.Value == true;
            byte readingOrder = ToReadingOrderCode(openXmlStyleIndex, alignment.ReadingOrder?.Value);
            bool hasAlignment = cellFormat?.ApplyAlignment?.Value == true
                || alignment.Horizontal != null
                || alignment.Vertical != null
                || alignment.WrapText != null
                || alignment.TextRotation != null
                || alignment.Indent != null
                || alignment.ShrinkToFit != null
                || alignment.ReadingOrder != null
                || (cellFormat?.Alignment == null && parentCellFormat?.ApplyAlignment?.Value == true);

            if (!hasAlignment) {
                return AlignmentInfo.None;
            }

            return new AlignmentInfo(horizontal, vertical, wrapText, textRotation, indent, shrinkToFit, readingOrder, hasAlignment);
        }

        private static byte ToHorizontalAlignmentCode(uint openXmlStyleIndex, HorizontalAlignmentValues? value) {
            if (!value.HasValue) return 0;
            if (value.Value == HorizontalAlignmentValues.General) return 0;
            if (value.Value == HorizontalAlignmentValues.Left) return 1;
            if (value.Value == HorizontalAlignmentValues.Center) return 2;
            if (value.Value == HorizontalAlignmentValues.Right) return 3;
            if (value.Value == HorizontalAlignmentValues.Fill) return 4;
            if (value.Value == HorizontalAlignmentValues.Justify) return 5;
            if (value.Value == HorizontalAlignmentValues.CenterContinuous) return 6;
            if (value.Value == HorizontalAlignmentValues.Distributed) return 7;
            throw new NotSupportedException($"Native XLS saving does not yet support horizontal alignment '{value}' for Open XML style index {openXmlStyleIndex}.");
        }

        private static byte ToVerticalAlignmentCode(uint openXmlStyleIndex, VerticalAlignmentValues? value) {
            if (!value.HasValue) return 0;
            if (value.Value == VerticalAlignmentValues.Top) return 0;
            if (value.Value == VerticalAlignmentValues.Center) return 1;
            if (value.Value == VerticalAlignmentValues.Bottom) return 2;
            if (value.Value == VerticalAlignmentValues.Justify) return 3;
            if (value.Value == VerticalAlignmentValues.Distributed) return 4;
            throw new NotSupportedException($"Native XLS saving does not yet support vertical alignment '{value}' for Open XML style index {openXmlStyleIndex}.");
        }

        private static byte ToTextRotationCode(uint openXmlStyleIndex, UInt32Value? value) {
            if (value == null) {
                return 0;
            }

            uint rotation = value.Value;
            if (rotation <= 180 || rotation == 255) {
                return checked((byte)rotation);
            }

            throw new NotSupportedException($"Native XLS saving does not yet support text rotation {rotation} for Open XML style index {openXmlStyleIndex}.");
        }

        private static byte ToIndentCode(uint openXmlStyleIndex, UInt32Value? value) {
            if (value == null) {
                return 0;
            }

            uint indent = value.Value;
            if (indent <= 15) {
                return checked((byte)indent);
            }

            throw new NotSupportedException($"Native XLS saving supports BIFF8 indentation levels up to 15; Open XML style index {openXmlStyleIndex} uses {indent}.");
        }

        private static byte ToReadingOrderCode(uint openXmlStyleIndex, UInt32Value? value) {
            if (value == null) {
                return 0;
            }

            uint readingOrder = value.Value;
            if (readingOrder <= 2) {
                return checked((byte)readingOrder);
            }

            throw new NotSupportedException($"Native XLS saving does not yet support reading order {readingOrder} for Open XML style index {openXmlStyleIndex}.");
        }

        private static void ThrowIfUnsupportedStyleFacet(uint openXmlStyleIndex, CellFormat? cellFormat) {
            if (cellFormat == null) {
                return;
            }

            bool unsupported = cellFormat.PivotButton?.Value == true;

            ThrowIfStyleExtensionsPresent(openXmlStyleIndex, cellFormat, "cell-format style");
            ThrowIfMultipleStyleChildren<Alignment>(openXmlStyleIndex, cellFormat, "cell-format style", "alignment");
            ThrowIfMultipleStyleChildren<Protection>(openXmlStyleIndex, cellFormat, "cell-format style", "protection");
            if (unsupported) {
                throw new NotSupportedException($"Native XLS saving currently supports number-format, basic font, solid/patterned fill, BIFF-backed alignment, basic border, cell-protection, and quote-prefix styles; Open XML style index {openXmlStyleIndex} uses unsupported style facets.");
            }
        }

        private static void ThrowIfStyleExtensionsPresent(uint openXmlStyleIndex, OpenXmlElement? element, string subject) {
            if (element?.GetFirstChild<ExtensionList>() != null) {
                throw new NotSupportedException($"Native XLS saving does not yet support {subject} extension payloads for Open XML style index {openXmlStyleIndex}.");
            }
        }

        private static void ThrowIfMultipleStyleChildren<TElement>(uint openXmlStyleIndex, OpenXmlElement? element, string subject, string elementName)
            where TElement : OpenXmlElement {
            if (element?.Elements<TElement>().Take(2).Count() > 1) {
                throw new NotSupportedException($"Native XLS saving does not yet support {subject} with duplicate {elementName} elements for Open XML style index {openXmlStyleIndex}.");
            }
        }

        private static ProtectionInfo ResolveProtection(CellFormat? cellFormat, CellFormat? parentCellFormat) {
            Protection? protection = cellFormat?.Protection ?? parentCellFormat?.Protection;
            bool quotePrefix = cellFormat?.QuotePrefix?.Value
                ?? parentCellFormat?.QuotePrefix?.Value
                ?? false;
            bool hasProtection = cellFormat?.ApplyProtection?.Value == true
                || protection != null
                || (cellFormat?.Protection == null && parentCellFormat?.ApplyProtection?.Value == true);
            if (!hasProtection && !quotePrefix) {
                return ProtectionInfo.Default;
            }

            bool locked = protection?.Locked?.Value ?? true;
            bool formulaHidden = protection?.Hidden?.Value == true;
            return new ProtectionInfo(locked, formulaHidden, quotePrefix, hasProtection);
        }

        private static BorderInfo ResolveBorder(
            uint openXmlStyleIndex,
            CellFormat? cellFormat,
            CellFormat? parentCellFormat,
            IReadOnlyList<Border> openXmlBorders,
            LegacyXlsFontTable fontTable) {
            UInt32Value? borderIdValue = cellFormat?.BorderId ?? parentCellFormat?.BorderId;
            if (borderIdValue == null || borderIdValue.Value == 0U) {
                return BorderInfo.None;
            }

            uint borderId = borderIdValue.Value;
            if (borderId >= openXmlBorders.Count) {
                throw new NotSupportedException($"Native XLS saving cannot resolve border id {borderId} for Open XML style index {openXmlStyleIndex}.");
            }

            Border border = openXmlBorders[(int)borderId];
            ThrowIfStyleExtensionsPresent(openXmlStyleIndex, border, "border style");
            ThrowIfMultipleStyleChildren<LeftBorder>(openXmlStyleIndex, border, "border style", "left border");
            ThrowIfMultipleStyleChildren<RightBorder>(openXmlStyleIndex, border, "border style", "right border");
            ThrowIfMultipleStyleChildren<TopBorder>(openXmlStyleIndex, border, "border style", "top border");
            ThrowIfMultipleStyleChildren<BottomBorder>(openXmlStyleIndex, border, "border style", "bottom border");
            ThrowIfMultipleStyleChildren<DiagonalBorder>(openXmlStyleIndex, border, "border style", "diagonal border");
            BorderSideInfo left = ResolveBorderSide(openXmlStyleIndex, border.LeftBorder, "left", fontTable);
            BorderSideInfo right = ResolveBorderSide(openXmlStyleIndex, border.RightBorder, "right", fontTable);
            BorderSideInfo top = ResolveBorderSide(openXmlStyleIndex, border.TopBorder, "top", fontTable);
            BorderSideInfo bottom = ResolveBorderSide(openXmlStyleIndex, border.BottomBorder, "bottom", fontTable);
            bool diagonalUp = border.DiagonalUp?.Value == true;
            bool diagonalDown = border.DiagonalDown?.Value == true;
            BorderSideInfo diagonal = diagonalUp || diagonalDown
                ? ResolveBorderSide(openXmlStyleIndex, border.DiagonalBorder, "diagonal", fontTable)
                : BorderSideInfo.None;

            bool hasBorder = left.HasBorder
                || right.HasBorder
                || top.HasBorder
                || bottom.HasBorder
                || diagonal.HasBorder;
            if (!hasBorder) {
                return BorderInfo.None;
            }

            return new BorderInfo(left, right, top, bottom, diagonal, diagonalUp, diagonalDown);
        }

        private static BorderSideInfo ResolveBorderSide(
            uint openXmlStyleIndex,
            BorderPropertiesType? side,
            string role,
            LegacyXlsFontTable fontTable) {
            byte style = ToBorderStyleCode(openXmlStyleIndex, side?.Style?.Value, role);
            if (style == 0) {
                return BorderSideInfo.None;
            }

            ThrowIfMultipleStyleChildren<Color>(openXmlStyleIndex, side, $"border {role} side", "color");
            ColorType? color = side?.GetFirstChild<Color>();
            if (!fontTable.TryGetColorIndex(color, $"cell border {role}", "colors", out ushort colorIndex, out string? reason)) {
                throw new NotSupportedException($"Native XLS saving does not yet support {reason ?? "cell border colors"} for Open XML style index {openXmlStyleIndex}.");
            }

            if (colorIndex == 0x7fff) {
                colorIndex = 0x0040;
            }

            if (colorIndex > 0x007f) {
                throw new NotSupportedException($"Native XLS saving does not yet support cell border {role} color indexes above 0x007F for Open XML style index {openXmlStyleIndex}.");
            }

            return new BorderSideInfo(style, colorIndex);
        }

        private static byte ToBorderStyleCode(uint openXmlStyleIndex, BorderStyleValues? value, string role) {
            if (!value.HasValue || value.Value == BorderStyleValues.None) return 0;
            if (value.Value == BorderStyleValues.Thin) return 1;
            if (value.Value == BorderStyleValues.Medium) return 2;
            if (value.Value == BorderStyleValues.Dashed) return 3;
            if (value.Value == BorderStyleValues.Dotted) return 4;
            if (value.Value == BorderStyleValues.Thick) return 5;
            if (value.Value == BorderStyleValues.Double) return 6;
            if (value.Value == BorderStyleValues.Hair) return 7;
            if (value.Value == BorderStyleValues.MediumDashed) return 8;
            if (value.Value == BorderStyleValues.DashDot) return 9;
            if (value.Value == BorderStyleValues.MediumDashDot) return 10;
            if (value.Value == BorderStyleValues.DashDotDot) return 11;
            if (value.Value == BorderStyleValues.MediumDashDotDot) return 12;
            if (value.Value == BorderStyleValues.SlantDashDot) return 13;
            throw new NotSupportedException($"Native XLS saving does not yet support {role} border style '{value}' for Open XML style index {openXmlStyleIndex}.");
        }

        private static byte[] BuildFormatPayload(ushort formatId, string formatCode) {
            byte[] formatBytes = EncodeUnicodeString(formatCode, out byte flags);
            if (formatCode.Length > ushort.MaxValue || 5L + formatBytes.Length > ushort.MaxValue) {
                throw new NotSupportedException($"Native XLS saving does not yet support custom number format lengths outside BIFF8 limits for format id {formatId}.");
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, formatId);
            WriteUInt16(stream, checked((ushort)formatCode.Length));
            stream.WriteByte(flags);
            stream.Write(formatBytes, 0, formatBytes.Length);
            return stream.ToArray();
        }

        private static byte[] BuildXfPayload(ushort numberFormatId, ushort fontIndex, bool applyFont, FillInfo fill, AlignmentInfo alignment, BorderInfo border, ProtectionInfo protection) {
            byte[] payload = new byte[20];
            WriteUInt16(payload, 0, fontIndex);
            WriteUInt16(payload, 2, numberFormatId);
            WriteUInt16(payload, 4, protection.Options);
            ushort attributes = 0;
            uint sideBits = 0;
            uint topBottomAndFillBits = 0;
            if (numberFormatId != 0) {
                attributes |= 0x0400;
            }

            if (applyFont) {
                attributes |= 0x0800;
            }

            if (fill.HasFill) {
                attributes |= 0x4000;
                topBottomAndFillBits |= (uint)fill.Pattern << 26;
                ushort fillColors = checked((ushort)((fill.BackgroundColorIndex << 7) | fill.ForegroundColorIndex));
                WriteUInt16(payload, 18, fillColors);
            } else {
                WriteUInt16(payload, 18, 0x2080);
            }

            if (border.HasBorder) {
                attributes |= 0x2000;
                sideBits |= border.Left.Style;
                sideBits |= (uint)border.Right.Style << 4;
                sideBits |= (uint)border.Top.Style << 8;
                sideBits |= (uint)border.Bottom.Style << 12;
                sideBits |= (uint)border.Left.ColorIndex << 16;
                sideBits |= (uint)border.Right.ColorIndex << 23;
                if (border.DiagonalDown) {
                    sideBits |= 1U << 30;
                }

                if (border.DiagonalUp) {
                    sideBits |= 1U << 31;
                }

                topBottomAndFillBits |= border.Top.ColorIndex;
                topBottomAndFillBits |= (uint)border.Bottom.ColorIndex << 7;
                topBottomAndFillBits |= (uint)border.Diagonal.ColorIndex << 14;
                topBottomAndFillBits |= (uint)border.Diagonal.Style << 21;
            }

            if (protection.HasProtection) {
                attributes |= 0x8000;
            }

            if (alignment.HasAlignment) {
                attributes |= 0x1000;
                payload[6] = checked((byte)(alignment.Horizontal | (alignment.WrapText ? 0x08 : 0) | (alignment.Vertical << 4)));
                payload[7] = alignment.TextRotation;
                attributes |= alignment.ExtendedOptions;
            }

            WriteUInt16(payload, 8, attributes);
            WriteUInt32(payload, 10, sideBits);
            WriteUInt32(payload, 14, topBottomAndFillBits);
            return payload;
        }

        private static bool FormatRecordMatches(byte[] record, ushort formatId, string formatCode) {
            byte[] expected = BuildFormatPayload(formatId, formatCode);
            return record.SequenceEqual(expected);
        }

        private static byte[] EncodeUnicodeString(string text, out byte flags) {
            if (CanUseCompressedString(text)) {
                flags = 0;
                return Encoding.ASCII.GetBytes(text);
            }

            flags = 1;
            return Encoding.Unicode.GetBytes(text);
        }

        private static bool CanUseCompressedString(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt16(byte[] buffer, int offset, ushort value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
        }

        private static void WriteUInt32(byte[] buffer, int offset, uint value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
            buffer[offset + 2] = (byte)((value >> 16) & 0xff);
            buffer[offset + 3] = (byte)((value >> 24) & 0xff);
        }

        private readonly struct FillInfo {
            internal static readonly FillInfo None = new FillInfo(0, 0, 0);

            internal FillInfo(byte pattern, ushort foregroundColorIndex, ushort backgroundColorIndex) {
                Pattern = pattern;
                ForegroundColorIndex = foregroundColorIndex;
                BackgroundColorIndex = backgroundColorIndex;
            }

            internal byte Pattern { get; }
            internal ushort ForegroundColorIndex { get; }
            internal ushort BackgroundColorIndex { get; }
            internal bool HasFill => Pattern != 0;
        }

        private readonly struct BorderSideInfo {
            internal static readonly BorderSideInfo None = new BorderSideInfo(0, 0);

            internal BorderSideInfo(byte style, ushort colorIndex) {
                Style = style;
                ColorIndex = colorIndex;
            }

            internal byte Style { get; }
            internal ushort ColorIndex { get; }
            internal bool HasBorder => Style != 0;
        }

        private readonly struct BorderInfo {
            internal static readonly BorderInfo None = new BorderInfo(BorderSideInfo.None, BorderSideInfo.None, BorderSideInfo.None, BorderSideInfo.None, BorderSideInfo.None, diagonalUp: false, diagonalDown: false);

            internal BorderInfo(BorderSideInfo left, BorderSideInfo right, BorderSideInfo top, BorderSideInfo bottom, BorderSideInfo diagonal, bool diagonalUp, bool diagonalDown) {
                Left = left;
                Right = right;
                Top = top;
                Bottom = bottom;
                Diagonal = diagonal;
                DiagonalUp = diagonalUp;
                DiagonalDown = diagonalDown;
            }

            internal BorderSideInfo Left { get; }
            internal BorderSideInfo Right { get; }
            internal BorderSideInfo Top { get; }
            internal BorderSideInfo Bottom { get; }
            internal BorderSideInfo Diagonal { get; }
            internal bool DiagonalUp { get; }
            internal bool DiagonalDown { get; }
            internal bool HasBorder => Left.HasBorder || Right.HasBorder || Top.HasBorder || Bottom.HasBorder || Diagonal.HasBorder;
        }

        private readonly struct ProtectionInfo {
            internal static readonly ProtectionInfo Default = new ProtectionInfo(locked: true, formulaHidden: false, quotePrefix: false, hasProtection: false);

            internal ProtectionInfo(bool locked, bool formulaHidden, bool quotePrefix, bool hasProtection) {
                Locked = locked;
                FormulaHidden = formulaHidden;
                QuotePrefix = quotePrefix;
                HasProtection = hasProtection;
            }

            internal bool Locked { get; }
            internal bool FormulaHidden { get; }
            internal bool QuotePrefix { get; }
            internal bool HasProtection { get; }
            internal ushort Options => checked((ushort)((Locked ? 0x0001 : 0) | (FormulaHidden ? 0x0002 : 0) | (QuotePrefix ? 0x0008 : 0)));
        }

        private readonly struct AlignmentInfo {
            internal static readonly AlignmentInfo None = new AlignmentInfo(0, 0, wrapText: false, textRotation: 0, indent: 0, shrinkToFit: false, readingOrder: 0, hasAlignment: false);

            internal AlignmentInfo(byte horizontal, byte vertical, bool wrapText, byte textRotation, byte indent, bool shrinkToFit, byte readingOrder, bool hasAlignment) {
                Horizontal = horizontal;
                Vertical = vertical;
                WrapText = wrapText;
                TextRotation = textRotation;
                Indent = indent;
                ShrinkToFit = shrinkToFit;
                ReadingOrder = readingOrder;
                HasAlignment = hasAlignment;
            }

            internal byte Horizontal { get; }
            internal byte Vertical { get; }
            internal bool WrapText { get; }
            internal byte TextRotation { get; }
            internal byte Indent { get; }
            internal bool ShrinkToFit { get; }
            internal byte ReadingOrder { get; }

            internal ushort ExtendedOptions => checked((ushort)(Indent | (ShrinkToFit ? 0x10 : 0) | (ReadingOrder << 6)));

            internal bool HasAlignment { get; }
        }
    }
}
