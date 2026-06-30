namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocParagraphFormattingReader {
        private const int OleSectorSize = 512;
        private const int PapxFkpBxLength = 13;
        private const ushort SprmPIstd = 0x4600;
        private const ushort SprmPFKeep = 0x2405;
        private const ushort SprmPFKeepFollow = 0x2406;
        private const ushort SprmPFPageBreakBefore = 0x2407;
        private const ushort SprmPFInTable = 0x2416;
        private const ushort SprmPFTtp = 0x2417;
        private const ushort SprmPJc = 0x2461;
        private const ushort SprmPJc80 = 0x2403;
        private const ushort SprmPDxaRight = 0x840E;
        private const ushort SprmPDxaLeft = 0x840F;
        private const ushort SprmPDxaLeft1 = 0x8411;
        private const ushort SprmPDyaLine = 0x6412;
        private const ushort SprmPDyaBefore = 0xA413;
        private const ushort SprmPDyaAfter = 0xA414;
        private const ushort SprmPBrcTop80 = 0x6424;
        private const ushort SprmPBrcLeft80 = 0x6425;
        private const ushort SprmPBrcBottom80 = 0x6426;
        private const ushort SprmPBrcRight80 = 0x6427;
        private const ushort SprmPBrcBetween80 = 0x6428;
        private const ushort SprmPShd80 = 0x442D;
        private const ushort SprmPFWidowControl = 0x2431;
        private const ushort SprmPChgTabsPapx = 0xC60D;
        private const ushort SprmPIlvl = 0x260A;
        private const ushort SprmPIlfo = 0x460B;
        private const ushort SprmTFCantSplit = 0x3403;
        private const ushort SprmTTableHeader = 0x3404;
        private const ushort SprmTFCantSplit90 = 0x3466;
        private const ushort SprmTJc = 0x548A;
        private const ushort SprmTDyaRowHeight = 0x9407;
        private const ushort SprmTFAutofit = 0x3615;
        private const ushort SprmTDefTable = 0xD608;
        private const ushort SprmTDefTableShd80 = 0xD609;
        private const ushort SprmTCellPadding = 0xD632;
        private const ushort SprmTCellSpacingDefault = 0xD633;
        private const ushort SprmTCellPaddingDefault = 0xD634;
        private const ushort SprmTTableWidth = 0xF614;
        private const ushort Shd80Nil = 0xFFFF;
        private const int Tc80Length = 20;
        private const byte FbrcTop = 0x01;
        private const byte FbrcLeft = 0x02;
        private const byte FbrcBottom = 0x04;
        private const byte FbrcRight = 0x08;
        private const byte FtsAuto = 0x01;
        private const byte FtsPercent = 0x02;
        private const byte FtsDxa = 0x03;
        private const ushort TcgrfHorizontalMergeMask = 0x0003;
        private const ushort TcgrfTextFlowMask = 0x001C;
        private const ushort TcgrfVerticalMergeMask = 0x0060;
        private const ushort TcgrfVerticalAlignmentMask = 0x0180;
        private const ushort TcgrfFitTextMask = 0x1000;
        private const ushort TcgrfNoWrapMask = 0x2000;
        private const ushort TcgrfHideMarkMask = 0x4000;

        internal static IReadOnlyList<LegacyDocParagraphFormatRange> ReadParagraphFormatting(
            byte[] wordDocumentStream,
            byte[] tableStream,
            LegacyDocFib fib,
            out string? warning) {
            warning = null;

            if (fib.LcbPlcfBtePapx == 0) {
                return Array.Empty<LegacyDocParagraphFormatRange>();
            }

            if (fib.FcPlcfBtePapx < 0
                || fib.LcbPlcfBtePapx < 4
                || fib.FcPlcfBtePapx + fib.LcbPlcfBtePapx > tableStream.Length
                || (fib.LcbPlcfBtePapx - 4) % 8 != 0) {
                warning = "The FIB points outside the selected table stream for the paragraph-format bin table.";
                return Array.Empty<LegacyDocParagraphFormatRange>();
            }

            int binCount = (fib.LcbPlcfBtePapx - 4) / 8;
            int cpArrayOffset = fib.FcPlcfBtePapx;
            int bteArrayOffset = cpArrayOffset + ((binCount + 1) * 4);
            var ranges = new List<LegacyDocParagraphFormatRange>();

            for (int binIndex = 0; binIndex < binCount; binIndex++) {
                int fcStart = LegacyDocFib.ReadInt32(tableStream, cpArrayOffset + (binIndex * 4));
                int fcEnd = LegacyDocFib.ReadInt32(tableStream, cpArrayOffset + ((binIndex + 1) * 4));
                int pageNumber = LegacyDocFib.ReadInt32(tableStream, bteArrayOffset + (binIndex * 4));
                if (fcEnd <= fcStart) {
                    continue;
                }

                int pageOffset = checked(pageNumber * OleSectorSize);
                if (pageOffset < 0 || pageOffset + OleSectorSize > wordDocumentStream.Length) {
                    warning = "A paragraph-format bin table entry points outside the WordDocument stream.";
                    return ranges;
                }

                ReadPapxFkp(wordDocumentStream, pageOffset, ranges);
            }

            return ranges
                .OrderBy(range => range.FileOffsetStart)
                .ThenBy(range => range.FileOffsetEnd)
                .ToArray();
        }

        private static void ReadPapxFkp(byte[] wordDocumentStream, int pageOffset, List<LegacyDocParagraphFormatRange> ranges) {
            int cpara = wordDocumentStream[pageOffset + OleSectorSize - 1];
            if (cpara <= 0) {
                return;
            }

            int rgfcOffset = pageOffset;
            int rgbxOffset = pageOffset + ((cpara + 1) * 4);
            if (rgbxOffset + (cpara * PapxFkpBxLength) > pageOffset + OleSectorSize - 1) {
                return;
            }

            for (int paragraphIndex = 0; paragraphIndex < cpara; paragraphIndex++) {
                int fcStart = LegacyDocFib.ReadInt32(wordDocumentStream, rgfcOffset + (paragraphIndex * 4));
                int fcEnd = LegacyDocFib.ReadInt32(wordDocumentStream, rgfcOffset + ((paragraphIndex + 1) * 4));
                if (fcEnd <= fcStart) {
                    continue;
                }

                int papxOffset = wordDocumentStream[rgbxOffset + (paragraphIndex * PapxFkpBxLength)] * 2;
                if (papxOffset == 0) {
                    continue;
                }

                int absolutePapxOffset = pageOffset + papxOffset;
                if (absolutePapxOffset >= pageOffset + OleSectorSize - 1) {
                    continue;
                }

                LegacyDocParagraphFormat format = ReadPapx(wordDocumentStream, absolutePapxOffset, pageOffset + OleSectorSize - 1);
                if (format.HasFormatting) {
                    ranges.Add(new LegacyDocParagraphFormatRange(fcStart, fcEnd, format));
                }
            }
        }

        private static LegacyDocParagraphFormat ReadPapx(byte[] bytes, int offset, int pageEnd) {
            if (offset >= pageEnd) {
                return LegacyDocParagraphFormat.Default;
            }

            int cb = bytes[offset];
            int grpprlOffset = offset + 1;
            int grpprlLength = cb * 2;
            if (cb == 0) {
                if (offset + 2 > pageEnd) {
                    return LegacyDocParagraphFormat.Default;
                }

                grpprlLength = bytes[offset + 1] * 2;
                grpprlOffset = offset + 2;
            }

            if (grpprlLength < 2 || grpprlOffset + grpprlLength > pageEnd) {
                return LegacyDocParagraphFormat.Default;
            }

            ushort styleIndex = LegacyDocFib.ReadUInt16(bytes, grpprlOffset);
            return ReadGrpprl(bytes, grpprlOffset + 2, grpprlLength - 2, styleIndex == 0 ? null : styleIndex);
        }

        internal static LegacyDocParagraphFormat ReadGrpprl(byte[] bytes, int offset, int count, ushort? baseStyleIndex = null) {
            int end = offset + count;
            LegacyDocParagraphAlignment? alignment = null;
            int? spacingBeforeTwips = null;
            int? spacingAfterTwips = null;
            int? lineSpacingTwips = null;
            int? leftIndentTwips = null;
            int? rightIndentTwips = null;
            int? firstLineIndentTwips = null;
            bool? keepLinesTogether = null;
            bool? keepWithNext = null;
            bool? pageBreakBefore = null;
            bool? avoidWidowAndOrphan = null;
            ushort? numberingListIndex = null;
            byte? numberingLevel = null;
            LegacyDocParagraphShading? paragraphShading = null;
            LegacyDocParagraphBorder paragraphTopBorder = default;
            LegacyDocParagraphBorder paragraphLeftBorder = default;
            LegacyDocParagraphBorder paragraphBottomBorder = default;
            LegacyDocParagraphBorder paragraphRightBorder = default;
            LegacyDocParagraphBorder paragraphBetweenBorder = default;
            bool? isInTable = null;
            bool? isTableTerminatingParagraph = null;
            var tabStops = new List<LegacyDocTabStop>();
            IReadOnlyList<int>? tableCellWidthsTwips = null;
            int? tableLeftIndentTwips = null;
            IReadOnlyList<LegacyDocTableCellHorizontalMerge>? tableCellHorizontalMerges = null;
            IReadOnlyList<LegacyDocTableCellVerticalMerge>? tableCellVerticalMerges = null;
            IReadOnlyList<LegacyDocTableCellVerticalAlignment>? tableCellVerticalAlignments = null;
            IReadOnlyList<LegacyDocTableCellTextDirection>? tableCellTextDirections = null;
            IReadOnlyList<bool>? tableCellFitTexts = null;
            IReadOnlyList<bool>? tableCellNoWraps = null;
            IReadOnlyList<bool>? tableCellHideMarks = null;
            IReadOnlyList<LegacyDocTableCellMargins>? tableCellMargins = null;
            IReadOnlyList<LegacyDocTableCellShading>? tableCellShadings = null;
            IReadOnlyList<LegacyDocTableCellBorders>? tableCellBorders = null;
            LegacyDocTableCellMargins? defaultTableCellMargins = null;
            int? defaultTableCellSpacingTwips = null;
            int? tableRowHeightTwips = null;
            bool tableRowHeightIsExact = false;
            bool? tableRowCantSplit = null;
            bool? tableRowIsHeader = null;
            LegacyDocTableAlignment? tableAlignment = null;
            LegacyDocTablePreferredWidth? tablePreferredWidth = null;
            bool? tableAutofit = null;
            bool hasMergedTableCells = false;
            ushort? styleIndex = baseStyleIndex;
            while (offset + 2 <= end) {
                ushort sprm = LegacyDocFib.ReadUInt16(bytes, offset);
                if (sprm == SprmPIstd) {
                    if (offset + 4 > end) {
                        break;
                    }

                    styleIndex = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    offset += 4;
                    continue;
                }

                if (sprm == SprmPFKeep
                    || sprm == SprmPFKeepFollow
                    || sprm == SprmPFPageBreakBefore
                    || sprm == SprmPFWidowControl
                    || sprm == SprmPFInTable
                    || sprm == SprmPFTtp) {
                    if (offset + 3 > end) {
                        break;
                    }

                    bool? value = ReadBoolOperand(bytes[offset + 2]);
                    switch (sprm) {
                        case SprmPFKeep:
                            keepLinesTogether = value;
                            break;
                        case SprmPFKeepFollow:
                            keepWithNext = value;
                            break;
                        case SprmPFPageBreakBefore:
                            pageBreakBefore = value;
                            break;
                        case SprmPFWidowControl:
                            avoidWidowAndOrphan = value;
                            break;
                        case SprmPFInTable:
                            isInTable = value;
                            break;
                        case SprmPFTtp:
                            isTableTerminatingParagraph = value;
                            break;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmPJc || sprm == SprmPJc80) {
                    if (offset + 3 > end) {
                        break;
                    }

                    alignment = MapAlignment(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmPIlvl) {
                    if (offset + 3 > end) {
                        break;
                    }

                    byte level = bytes[offset + 2];
                    if (level <= 8) {
                        numberingLevel = level;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmPIlfo) {
                    if (offset + 4 > end) {
                        break;
                    }

                    ushort ilfo = LegacyDocFib.ReadUInt16(bytes, offset + 2);
                    if (ilfo > 0) {
                        numberingListIndex = ilfo;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmTFCantSplit || sprm == SprmTFCantSplit90 || sprm == SprmTTableHeader) {
                    if (offset + 3 > end) {
                        break;
                    }

                    bool? value = ReadBoolOperand(bytes[offset + 2]);
                    if (sprm == SprmTTableHeader) {
                        tableRowIsHeader = value;
                    } else {
                        tableRowCantSplit = value;
                    }

                    offset += 3;
                    continue;
                }

                if (sprm == SprmTFAutofit) {
                    if (offset + 3 > end) {
                        break;
                    }

                    tableAutofit = bytes[offset + 2] != 0;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmTJc) {
                    if (offset + 4 > end) {
                        break;
                    }

                    tableAlignment = MapTableAlignment(LegacyDocFib.ReadUInt16(bytes, offset + 2));
                    offset += 4;
                    continue;
                }

                if (sprm == SprmTTableWidth) {
                    if (offset + 5 > end) {
                        break;
                    }

                    tablePreferredWidth = ReadTablePreferredWidth(bytes[offset + 2], LegacyDocFib.ReadUInt16(bytes, offset + 3));
                    offset += 5;
                    continue;
                }

                if (sprm == SprmPDxaLeft || sprm == SprmPDxaRight || sprm == SprmPDxaLeft1 || sprm == SprmPDyaBefore || sprm == SprmPDyaAfter) {
                    if (offset + 4 > end) {
                        break;
                    }

                    int value = ReadInt16(bytes, offset + 2);
                    switch (sprm) {
                        case SprmPDxaLeft:
                            leftIndentTwips = value;
                            break;
                        case SprmPDxaRight:
                            rightIndentTwips = value;
                            break;
                        case SprmPDxaLeft1:
                            firstLineIndentTwips = value;
                            break;
                        case SprmPDyaBefore:
                            spacingBeforeTwips = value;
                            break;
                        case SprmPDyaAfter:
                            spacingAfterTwips = value;
                            break;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmPDyaLine) {
                    if (offset + 6 > end) {
                        break;
                    }

                    int dyaLine = ReadInt16(bytes, offset + 2);
                    int fMultLinespace = ReadInt16(bytes, offset + 4);
                    if (fMultLinespace == 0 && dyaLine > 0) {
                        lineSpacingTwips = dyaLine;
                    }

                    offset += 6;
                    continue;
                }

                if (sprm == SprmPShd80) {
                    if (offset + 4 > end) {
                        break;
                    }

                    paragraphShading = ReadParagraphShading(LegacyDocFib.ReadUInt16(bytes, offset + 2));
                    offset += 4;
                    continue;
                }

                if (sprm == SprmPBrcTop80
                    || sprm == SprmPBrcLeft80
                    || sprm == SprmPBrcBottom80
                    || sprm == SprmPBrcRight80
                    || sprm == SprmPBrcBetween80) {
                    if (offset + 6 > end) {
                        break;
                    }

                    LegacyDocParagraphBorder border = ReadParagraphBorder(bytes, offset + 2);
                    switch (sprm) {
                        case SprmPBrcTop80:
                            paragraphTopBorder = border;
                            break;
                        case SprmPBrcLeft80:
                            paragraphLeftBorder = border;
                            break;
                        case SprmPBrcBottom80:
                            paragraphBottomBorder = border;
                            break;
                        case SprmPBrcRight80:
                            paragraphRightBorder = border;
                            break;
                        case SprmPBrcBetween80:
                            paragraphBetweenBorder = border;
                            break;
                    }

                    offset += 6;
                    continue;
                }

                if (sprm == SprmPChgTabsPapx) {
                    if (offset + 3 > end) {
                        break;
                    }

                    int tabOperandLength = bytes[offset + 2];
                    if (offset + 3 + tabOperandLength > end) {
                        break;
                    }

                    ReadTabChanges(bytes, offset + 3, offset + 3 + tabOperandLength, tabStops);
                    offset += 3 + tabOperandLength;
                    continue;
                }

                if (sprm == SprmTDyaRowHeight) {
                    if (offset + 4 > end) {
                        break;
                    }

                    int rowHeight = ReadInt16(bytes, offset + 2);
                    if (rowHeight < 0) {
                        tableRowHeightTwips = -rowHeight;
                        tableRowHeightIsExact = true;
                    } else if (rowHeight > 0) {
                        tableRowHeightTwips = rowHeight;
                        tableRowHeightIsExact = false;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmTDefTable) {
                    if (!TryReadTableDefinition(
                        bytes,
                        offset,
                        end,
                        out tableCellWidthsTwips,
                        out tableLeftIndentTwips,
                        out tableCellHorizontalMerges,
                        out tableCellVerticalMerges,
                        out tableCellVerticalAlignments,
                        out tableCellTextDirections,
                        out tableCellFitTexts,
                        out tableCellNoWraps,
                        out tableCellHideMarks,
                        out tableCellBorders,
                        out bool tableDefinitionHasUnsupportedMergedCells,
                        out int tableDefinitionOperandLength)) {
                        break;
                    }

                    hasMergedTableCells |= tableDefinitionHasUnsupportedMergedCells;
                    offset += 2 + tableDefinitionOperandLength;
                    continue;
                }

                if (sprm == SprmTDefTableShd80) {
                    if (!TryReadTableCellShadings(
                        bytes,
                        offset,
                        end,
                        out tableCellShadings,
                        out int tableCellShadingOperandLength)) {
                        break;
                    }

                    offset += 2 + tableCellShadingOperandLength;
                    continue;
                }

                if (sprm == SprmTCellPadding || sprm == SprmTCellPaddingDefault) {
                    if (!TryReadTableCellPadding(
                        bytes,
                        offset,
                        end,
                        sprm == SprmTCellPaddingDefault,
                        ref tableCellMargins,
                        ref defaultTableCellMargins,
                        out int tableCellPaddingOperandLength)) {
                        break;
                    }

                    offset += 2 + tableCellPaddingOperandLength;
                    continue;
                }

                if (sprm == SprmTCellSpacingDefault) {
                    if (!TryReadTableCellSpacing(
                        bytes,
                        offset,
                        end,
                        ref defaultTableCellSpacingTwips,
                        out int tableCellSpacingOperandLength)) {
                        break;
                    }

                    offset += 2 + tableCellSpacingOperandLength;
                    continue;
                }

                if (!TryGetSprmOperandLength(bytes, offset, end, out int operandLength)) {
                    break;
                }

                offset += 2 + operandLength;
            }

            return new LegacyDocParagraphFormat(
                alignment,
                styleIndex,
                spacingBeforeTwips,
                spacingAfterTwips,
                lineSpacingTwips,
                leftIndentTwips,
                rightIndentTwips,
                firstLineIndentTwips,
                keepLinesTogether,
                keepWithNext,
                pageBreakBefore,
                avoidWidowAndOrphan,
                numberingListIndex,
                numberingLevel,
                isInTable,
                isTableTerminatingParagraph,
                tabStops,
                tableCellWidthsTwips,
                tableLeftIndentTwips,
                tableRowHeightTwips,
                tableRowHeightIsExact,
                tableRowCantSplit,
                tableRowIsHeader,
                tableAlignment,
                tablePreferredWidth,
                tableAutofit,
                tableCellHorizontalMerges,
                tableCellVerticalMerges,
                tableCellVerticalAlignments,
                tableCellTextDirections,
                tableCellFitTexts,
                tableCellNoWraps,
                tableCellHideMarks,
                tableCellMargins,
                tableCellShadings,
                tableCellBorders,
                defaultTableCellMargins,
                defaultTableCellSpacingTwips,
                hasMergedTableCells,
                paragraphShading,
                new LegacyDocParagraphBorders(
                    paragraphTopBorder,
                    paragraphLeftBorder,
                    paragraphBottomBorder,
                    paragraphRightBorder,
                    paragraphBetweenBorder));
        }

        private static LegacyDocParagraphShading ReadParagraphShading(ushort shd80) {
            if (shd80 == 0 || shd80 == Shd80Nil) {
                return default;
            }

            byte backgroundIco = (byte)((shd80 >> 5) & 0x1F);
            string? fillColorHex = LegacyDocColorPalette.GetHexForIco(backgroundIco);
            return string.IsNullOrEmpty(fillColorHex)
                ? default
                : new LegacyDocParagraphShading(fillColorHex);
        }

        private static LegacyDocParagraphBorder ReadParagraphBorder(byte[] bytes, int offset) {
            if (offset + 4 > bytes.Length) {
                return default;
            }

            if (bytes[offset] == 0xFF
                && bytes[offset + 1] == 0xFF
                && bytes[offset + 2] == 0xFF
                && bytes[offset + 3] == 0xFF) {
                return default;
            }

            byte sizeEighthPoints = bytes[offset];
            byte borderType = bytes[offset + 1];
            byte colorIndex = bytes[offset + 2];
            byte spacePoints = bytes[offset + 3];
            LegacyDocParagraphBorderStyle style = MapParagraphBorderStyle(borderType);
            if (style == LegacyDocParagraphBorderStyle.None) {
                return default;
            }

            string? colorHex = LegacyDocColorPalette.GetHexForIco(colorIndex);
            return new LegacyDocParagraphBorder(style, colorHex, sizeEighthPoints, spacePoints);
        }

        private static LegacyDocParagraphBorderStyle MapParagraphBorderStyle(byte borderType) {
            switch (borderType) {
                case 0x01:
                    return LegacyDocParagraphBorderStyle.Single;
                case 0x03:
                    return LegacyDocParagraphBorderStyle.Double;
                case 0x06:
                    return LegacyDocParagraphBorderStyle.Dotted;
                case 0x07:
                    return LegacyDocParagraphBorderStyle.Dashed;
                default:
                    return LegacyDocParagraphBorderStyle.None;
            }
        }

        private static bool TryReadTableDefinition(
            byte[] bytes,
            int sprmOffset,
            int end,
            out IReadOnlyList<int>? tableCellWidthsTwips,
            out int? tableLeftIndentTwips,
            out IReadOnlyList<LegacyDocTableCellHorizontalMerge>? tableCellHorizontalMerges,
            out IReadOnlyList<LegacyDocTableCellVerticalMerge>? tableCellVerticalMerges,
            out IReadOnlyList<LegacyDocTableCellVerticalAlignment>? tableCellVerticalAlignments,
            out IReadOnlyList<LegacyDocTableCellTextDirection>? tableCellTextDirections,
            out IReadOnlyList<bool>? tableCellFitTexts,
            out IReadOnlyList<bool>? tableCellNoWraps,
            out IReadOnlyList<bool>? tableCellHideMarks,
            out IReadOnlyList<LegacyDocTableCellBorders>? tableCellBorders,
            out bool hasUnsupportedMergedTableCells,
            out int operandLength) {
            tableCellWidthsTwips = null;
            tableLeftIndentTwips = null;
            tableCellHorizontalMerges = null;
            tableCellVerticalMerges = null;
            tableCellVerticalAlignments = null;
            tableCellTextDirections = null;
            tableCellFitTexts = null;
            tableCellNoWraps = null;
            tableCellHideMarks = null;
            tableCellBorders = null;
            hasUnsupportedMergedTableCells = false;
            operandLength = 0;
            if (sprmOffset + 5 > end) {
                return false;
            }

            ushort cb = LegacyDocFib.ReadUInt16(bytes, sprmOffset + 2);
            operandLength = cb + 1;
            int operandOffset = sprmOffset + 2;
            int operandEnd = operandOffset + operandLength;
            if (operandEnd > end || cb < 4) {
                return false;
            }

            int columnCount = bytes[sprmOffset + 4];
            if (columnCount <= 0) {
                tableCellWidthsTwips = Array.Empty<int>();
                tableLeftIndentTwips = null;
                tableCellHorizontalMerges = Array.Empty<LegacyDocTableCellHorizontalMerge>();
                tableCellVerticalMerges = Array.Empty<LegacyDocTableCellVerticalMerge>();
                tableCellVerticalAlignments = Array.Empty<LegacyDocTableCellVerticalAlignment>();
                tableCellTextDirections = Array.Empty<LegacyDocTableCellTextDirection>();
                tableCellFitTexts = Array.Empty<bool>();
                tableCellNoWraps = Array.Empty<bool>();
                tableCellHideMarks = Array.Empty<bool>();
                tableCellBorders = Array.Empty<LegacyDocTableCellBorders>();
                return true;
            }

            int edgesOffset = sprmOffset + 5;
            int tc80Offset = edgesOffset + ((columnCount + 1) * 2);
            if (tc80Offset + (columnCount * Tc80Length) > operandEnd) {
                return false;
            }

            var widths = new int[columnCount];
            var horizontalMerges = new LegacyDocTableCellHorizontalMerge[columnCount];
            var verticalMerges = new LegacyDocTableCellVerticalMerge[columnCount];
            var verticalAlignments = new LegacyDocTableCellVerticalAlignment[columnCount];
            var textDirections = new LegacyDocTableCellTextDirection[columnCount];
            var fitTexts = new bool[columnCount];
            var noWraps = new bool[columnCount];
            var hideMarks = new bool[columnCount];
            var borders = new LegacyDocTableCellBorders[columnCount];
            int previousEdge = ReadInt16(bytes, edgesOffset);
            if (previousEdge > 0) {
                tableLeftIndentTwips = previousEdge;
            }

            for (int index = 0; index < columnCount; index++) {
                int nextEdge = ReadInt16(bytes, edgesOffset + ((index + 1) * 2));
                int width = nextEdge - previousEdge;
                widths[index] = width > 0 ? width : 0;
                previousEdge = nextEdge;
            }

            for (int index = 0; index < columnCount; index++) {
                ushort tcgrf = LegacyDocFib.ReadUInt16(bytes, tc80Offset + (index * Tc80Length));
                switch (tcgrf & TcgrfHorizontalMergeMask) {
                    case 0:
                        horizontalMerges[index] = LegacyDocTableCellHorizontalMerge.None;
                        break;
                    case 0x0001:
                        horizontalMerges[index] = LegacyDocTableCellHorizontalMerge.Restart;
                        break;
                    case 0x0002:
                        horizontalMerges[index] = LegacyDocTableCellHorizontalMerge.Continue;
                        break;
                    default:
                        horizontalMerges[index] = LegacyDocTableCellHorizontalMerge.None;
                        hasUnsupportedMergedTableCells = true;
                        break;
                }

                switch (tcgrf & TcgrfVerticalMergeMask) {
                    case 0:
                        verticalMerges[index] = LegacyDocTableCellVerticalMerge.None;
                        break;
                    case 0x0020:
                        verticalMerges[index] = LegacyDocTableCellVerticalMerge.Restart;
                        break;
                    case 0x0040:
                        verticalMerges[index] = LegacyDocTableCellVerticalMerge.Continue;
                        break;
                    default:
                        verticalMerges[index] = LegacyDocTableCellVerticalMerge.None;
                        hasUnsupportedMergedTableCells = true;
                        break;
                }

                switch (tcgrf & TcgrfVerticalAlignmentMask) {
                    case 0:
                        verticalAlignments[index] = LegacyDocTableCellVerticalAlignment.Top;
                        break;
                    case 0x0080:
                        verticalAlignments[index] = LegacyDocTableCellVerticalAlignment.Center;
                        break;
                    case 0x0100:
                        verticalAlignments[index] = LegacyDocTableCellVerticalAlignment.Bottom;
                        break;
                    default:
                        verticalAlignments[index] = LegacyDocTableCellVerticalAlignment.Top;
                        break;
                }

                switch ((tcgrf & TcgrfTextFlowMask) >> 2) {
                    case 0:
                        textDirections[index] = LegacyDocTableCellTextDirection.LeftToRightTopToBottom;
                        break;
                    case 1:
                        textDirections[index] = LegacyDocTableCellTextDirection.TopToBottomRightToLeft;
                        break;
                    case 3:
                        textDirections[index] = LegacyDocTableCellTextDirection.BottomToTopLeftToRight;
                        break;
                    case 4:
                        textDirections[index] = LegacyDocTableCellTextDirection.LeftToRightTopToBottomRotated;
                        break;
                    case 5:
                        textDirections[index] = LegacyDocTableCellTextDirection.TopToBottomRightToLeftRotated;
                        break;
                    default:
                        textDirections[index] = LegacyDocTableCellTextDirection.LeftToRightTopToBottom;
                        break;
                }

                fitTexts[index] = (tcgrf & TcgrfFitTextMask) != 0;
                noWraps[index] = (tcgrf & TcgrfNoWrapMask) != 0;
                hideMarks[index] = (tcgrf & TcgrfHideMarkMask) != 0;
                borders[index] = ReadTableCellBorders(bytes, tc80Offset + (index * Tc80Length));
            }

            tableCellWidthsTwips = widths.Where(width => width > 0).ToArray();
            tableCellHorizontalMerges = horizontalMerges.Any(merge => merge != LegacyDocTableCellHorizontalMerge.None)
                ? horizontalMerges
                : Array.Empty<LegacyDocTableCellHorizontalMerge>();
            tableCellVerticalMerges = verticalMerges.Any(merge => merge != LegacyDocTableCellVerticalMerge.None)
                ? verticalMerges
                : Array.Empty<LegacyDocTableCellVerticalMerge>();
            tableCellVerticalAlignments = verticalAlignments.Any(alignment => alignment != LegacyDocTableCellVerticalAlignment.Top)
                ? verticalAlignments
                : Array.Empty<LegacyDocTableCellVerticalAlignment>();
            tableCellTextDirections = textDirections.Any(textDirection => textDirection != LegacyDocTableCellTextDirection.LeftToRightTopToBottom)
                ? textDirections
                : Array.Empty<LegacyDocTableCellTextDirection>();
            tableCellFitTexts = fitTexts.Any(fitText => fitText)
                ? fitTexts
                : Array.Empty<bool>();
            tableCellNoWraps = noWraps.Any(noWrap => noWrap)
                ? noWraps
                : Array.Empty<bool>();
            tableCellHideMarks = hideMarks.Any(hideMark => hideMark)
                ? hideMarks
                : Array.Empty<bool>();
            tableCellBorders = borders.Any(border => border.HasAny)
                ? borders
                : Array.Empty<LegacyDocTableCellBorders>();
            return true;
        }

        private static LegacyDocTableCellBorders ReadTableCellBorders(byte[] bytes, int tc80Offset) {
            return new LegacyDocTableCellBorders(
                ReadBrc80(bytes, tc80Offset + 4),
                ReadBrc80(bytes, tc80Offset + 8),
                ReadBrc80(bytes, tc80Offset + 12),
                ReadBrc80(bytes, tc80Offset + 16));
        }

        private static LegacyDocTableCellBorder ReadBrc80(byte[] bytes, int offset) {
            if (offset + 4 > bytes.Length) {
                return default;
            }

            if (bytes[offset] == 0xFF
                && bytes[offset + 1] == 0xFF
                && bytes[offset + 2] == 0xFF
                && bytes[offset + 3] == 0xFF) {
                return default;
            }

            byte sizeEighthPoints = bytes[offset];
            byte borderType = bytes[offset + 1];
            byte colorIndex = bytes[offset + 2];
            byte spacePoints = bytes[offset + 3];
            LegacyDocTableCellBorderStyle style = MapBrc80BorderStyle(borderType);
            if (style == LegacyDocTableCellBorderStyle.None) {
                return default;
            }

            string? colorHex = LegacyDocColorPalette.GetHexForIco(colorIndex);
            return new LegacyDocTableCellBorder(style, colorHex, sizeEighthPoints, spacePoints);
        }

        private static LegacyDocTableCellBorderStyle MapBrc80BorderStyle(byte borderType) {
            switch (borderType) {
                case 0x01:
                    return LegacyDocTableCellBorderStyle.Single;
                case 0x03:
                    return LegacyDocTableCellBorderStyle.Double;
                case 0x06:
                    return LegacyDocTableCellBorderStyle.Dotted;
                case 0x07:
                    return LegacyDocTableCellBorderStyle.Dashed;
                default:
                    return LegacyDocTableCellBorderStyle.None;
            }
        }

        private static bool TryReadTableCellPadding(
            byte[] bytes,
            int sprmOffset,
            int end,
            bool isDefault,
            ref IReadOnlyList<LegacyDocTableCellMargins>? tableCellMargins,
            ref LegacyDocTableCellMargins? defaultTableCellMargins,
            out int operandLength) {
            operandLength = 0;
            if (sprmOffset + 9 > end) {
                return false;
            }

            int cb = bytes[sprmOffset + 2];
            operandLength = 1 + cb;
            if (cb != 6 || sprmOffset + 2 + operandLength > end) {
                return false;
            }

            int itcFirst = bytes[sprmOffset + 3];
            int itcLim = bytes[sprmOffset + 4];
            byte grfbrc = bytes[sprmOffset + 5];
            byte ftsWidth = bytes[sprmOffset + 6];
            int width = LegacyDocFib.ReadUInt16(bytes, sprmOffset + 7);
            if (ftsWidth != FtsDxa || width < 0 || width > 31680) {
                return true;
            }

            LegacyDocTableCellMargins margins = CreateTableCellMargins(grfbrc, width);
            if (!margins.HasAny) {
                return true;
            }

            if (isDefault) {
                defaultTableCellMargins = (defaultTableCellMargins ?? default).Merge(margins);
                return true;
            }

            if (itcFirst >= itcLim) {
                return true;
            }

            LegacyDocTableCellMargins[] marginArray;
            if (tableCellMargins == null || tableCellMargins.Count < itcLim) {
                marginArray = new LegacyDocTableCellMargins[itcLim];
                if (tableCellMargins != null) {
                    for (int index = 0; index < tableCellMargins.Count; index++) {
                        marginArray[index] = tableCellMargins[index];
                    }
                }
            } else {
                marginArray = tableCellMargins.ToArray();
            }
            for (int index = itcFirst; index < itcLim; index++) {
                marginArray[index] = marginArray[index].Merge(margins);
            }

            tableCellMargins = marginArray.Any(margin => margin.HasAny)
                ? marginArray
                : Array.Empty<LegacyDocTableCellMargins>();
            return true;
        }

        private static bool TryReadTableCellSpacing(
            byte[] bytes,
            int sprmOffset,
            int end,
            ref int? defaultTableCellSpacingTwips,
            out int operandLength) {
            operandLength = 0;
            if (sprmOffset + 9 > end) {
                return false;
            }

            int cb = bytes[sprmOffset + 2];
            operandLength = 1 + cb;
            if (cb != 6 || sprmOffset + 2 + operandLength > end) {
                return false;
            }

            byte ftsWidth = bytes[sprmOffset + 6];
            int width = LegacyDocFib.ReadUInt16(bytes, sprmOffset + 7);
            if (ftsWidth == FtsDxa && width >= 0 && width <= 31680) {
                defaultTableCellSpacingTwips = width;
            }

            return true;
        }

        private static bool TryReadTableCellShadings(
            byte[] bytes,
            int sprmOffset,
            int end,
            out IReadOnlyList<LegacyDocTableCellShading>? tableCellShadings,
            out int operandLength) {
            tableCellShadings = null;
            operandLength = 0;
            if (sprmOffset + 3 > end) {
                return false;
            }

            int cb = bytes[sprmOffset + 2];
            operandLength = 1 + cb;
            if (cb % 2 != 0 || sprmOffset + 2 + operandLength > end) {
                return false;
            }

            int cellCount = cb / 2;
            var shadings = new LegacyDocTableCellShading[cellCount];
            for (int index = 0; index < cellCount; index++) {
                ushort shd80 = LegacyDocFib.ReadUInt16(bytes, sprmOffset + 3 + (index * 2));
                shadings[index] = ReadTableCellShading(shd80);
            }

            tableCellShadings = shadings.Any(shading => shading.HasAny)
                ? shadings
                : Array.Empty<LegacyDocTableCellShading>();
            return true;
        }

        private static LegacyDocTableCellShading ReadTableCellShading(ushort shd80) {
            if (shd80 == 0 || shd80 == Shd80Nil) {
                return default;
            }

            byte backgroundIco = (byte)((shd80 >> 5) & 0x1F);
            string? fillColorHex = LegacyDocColorPalette.GetHexForIco(backgroundIco);
            return string.IsNullOrEmpty(fillColorHex)
                ? default
                : new LegacyDocTableCellShading(fillColorHex);
        }

        private static LegacyDocTableCellMargins CreateTableCellMargins(byte sideMask, int widthTwips) {
            return new LegacyDocTableCellMargins(
                (sideMask & FbrcTop) != 0 ? widthTwips : null,
                (sideMask & FbrcRight) != 0 ? widthTwips : null,
                (sideMask & FbrcBottom) != 0 ? widthTwips : null,
                (sideMask & FbrcLeft) != 0 ? widthTwips : null);
        }

        private static void ReadTabChanges(byte[] bytes, int offset, int end, List<LegacyDocTabStop> tabStops) {
            if (offset >= end) {
                return;
            }

            int deletedCount = bytes[offset++];
            if (offset + (deletedCount * 2) > end) {
                return;
            }

            for (int index = 0; index < deletedCount; index++) {
                tabStops.Add(new LegacyDocTabStop(ReadInt16(bytes, offset), LegacyDocTabStopAlignment.Clear, LegacyDocTabStopLeader.None));
                offset += 2;
            }

            if (offset >= end) {
                return;
            }

            int addedCount = bytes[offset++];
            int addedPositionsOffset = offset;
            int addedDescriptorsOffset = addedPositionsOffset + (addedCount * 2);
            if (addedDescriptorsOffset + addedCount > end) {
                return;
            }

            for (int index = 0; index < addedCount; index++) {
                int position = ReadInt16(bytes, addedPositionsOffset + (index * 2));
                byte descriptor = bytes[addedDescriptorsOffset + index];
                if (TryMapTabAlignment((byte)(descriptor & 0x07), out LegacyDocTabStopAlignment alignment)
                    && TryMapTabLeader((byte)((descriptor >> 3) & 0x07), out LegacyDocTabStopLeader leader)) {
                    tabStops.Add(new LegacyDocTabStop(position, alignment, leader));
                }
            }
        }

        private static bool TryMapTabAlignment(byte value, out LegacyDocTabStopAlignment alignment) {
            switch (value) {
                case 0:
                    alignment = LegacyDocTabStopAlignment.Left;
                    return true;
                case 1:
                    alignment = LegacyDocTabStopAlignment.Center;
                    return true;
                case 2:
                    alignment = LegacyDocTabStopAlignment.Right;
                    return true;
                case 3:
                    alignment = LegacyDocTabStopAlignment.Decimal;
                    return true;
                case 4:
                    alignment = LegacyDocTabStopAlignment.Bar;
                    return true;
                default:
                    alignment = LegacyDocTabStopAlignment.Left;
                    return false;
            }
        }

        private static bool TryMapTabLeader(byte value, out LegacyDocTabStopLeader leader) {
            switch (value) {
                case 0:
                    leader = LegacyDocTabStopLeader.None;
                    return true;
                case 1:
                    leader = LegacyDocTabStopLeader.Dot;
                    return true;
                case 2:
                    leader = LegacyDocTabStopLeader.Hyphen;
                    return true;
                case 3:
                    leader = LegacyDocTabStopLeader.Underscore;
                    return true;
                case 4:
                    leader = LegacyDocTabStopLeader.Heavy;
                    return true;
                case 5:
                    leader = LegacyDocTabStopLeader.MiddleDot;
                    return true;
                default:
                    leader = LegacyDocTabStopLeader.None;
                    return false;
            }
        }

        private static bool? ReadBoolOperand(byte value) {
            return value == 0 ? null : true;
        }

        private static LegacyDocParagraphAlignment? MapAlignment(byte value) {
            switch (value) {
                case 0:
                    return LegacyDocParagraphAlignment.Left;
                case 1:
                    return LegacyDocParagraphAlignment.Center;
                case 2:
                    return LegacyDocParagraphAlignment.Right;
                case 3:
                    return LegacyDocParagraphAlignment.Justify;
                default:
                    return null;
            }
        }

        private static LegacyDocTableAlignment? MapTableAlignment(ushort value) {
            switch (value) {
                case 0:
                    return LegacyDocTableAlignment.Left;
                case 1:
                    return LegacyDocTableAlignment.Center;
                case 2:
                    return LegacyDocTableAlignment.Right;
                default:
                    return null;
            }
        }

        private static LegacyDocTablePreferredWidth? ReadTablePreferredWidth(byte ftsWidth, ushort width) {
            switch (ftsWidth) {
                case FtsAuto:
                    return new LegacyDocTablePreferredWidth(LegacyDocTablePreferredWidthUnit.Auto, 0);
                case FtsPercent:
                    return width <= short.MaxValue
                        ? new LegacyDocTablePreferredWidth(LegacyDocTablePreferredWidthUnit.Percent, width)
                        : (LegacyDocTablePreferredWidth?)null;
                case FtsDxa:
                    return width <= short.MaxValue
                        ? new LegacyDocTablePreferredWidth(LegacyDocTablePreferredWidthUnit.Dxa, width)
                        : (LegacyDocTablePreferredWidth?)null;
                default:
                    return null;
            }
        }

        private static bool TryGetSprmOperandLength(byte[] bytes, int sprmOffset, int end, out int operandLength) {
            operandLength = 0;
            ushort sprm = LegacyDocFib.ReadUInt16(bytes, sprmOffset);
            int spra = (sprm >> 13) & 0x7;
            switch (spra) {
                case 0:
                case 1:
                    operandLength = 1;
                    return sprmOffset + 2 + operandLength <= end;
                case 2:
                case 4:
                case 5:
                    operandLength = 2;
                    return sprmOffset + 2 + operandLength <= end;
                case 3:
                    operandLength = 4;
                    return sprmOffset + 2 + operandLength <= end;
                case 6:
                    if (sprmOffset + 3 > end) {
                        return false;
                    }

                    operandLength = 1 + bytes[sprmOffset + 2];
                    return sprmOffset + 2 + operandLength <= end;
                case 7:
                    operandLength = 3;
                    return sprmOffset + 2 + operandLength <= end;
                default:
                    return false;
            }
        }

        private static short ReadInt16(byte[] bytes, int offset) {
            return unchecked((short)LegacyDocFib.ReadUInt16(bytes, offset));
        }
    }
}
