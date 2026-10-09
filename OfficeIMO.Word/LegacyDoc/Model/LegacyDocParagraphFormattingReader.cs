namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static partial class LegacyDocParagraphFormattingReader {
        private const int OleSectorSize = 512;
        private const int PapxFkpBxLength = 13;
        private const ushort SprmPIstd = 0x4600;
        private const ushort SprmPFKeep = 0x2405;
        private const ushort SprmPFKeepFollow = 0x2406;
        private const ushort SprmPFPageBreakBefore = 0x2407;
        private const ushort SprmPFNoLineNumb = 0x240C;
        private const ushort SprmPFNoAutoHyph = 0x242A;
        private const ushort SprmPFKinsoku = 0x2433;
        private const ushort SprmPFWordWrap = 0x2434;
        private const ushort SprmPFOverflowPunct = 0x2435;
        private const ushort SprmPFTopLinePunct = 0x2436;
        private const ushort SprmPFAutoSpaceDE = 0x2437;
        private const ushort SprmPFAutoSpaceDN = 0x2438;
        private const ushort SprmPFContextualSpacing = 0x246D;
        private const ushort SprmPFMirrorIndents = 0x2470;
        private const ushort SprmPFInTable = 0x2416;
        private const ushort SprmPFTtp = 0x2417;
        private const ushort SprmPItap = 0x6649;
        private const ushort SprmPDtap = 0x664A;
        private const ushort SprmPFInnerTableCell = 0x244B;
        private const ushort SprmPFInnerTtp = 0x244C;
        private const ushort SprmPJc = 0x2461;
        private const ushort SprmPJc80 = 0x2403;
        private const ushort SprmPFBiDi = 0x2441;
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
        private const ushort SprmPChgTabs = 0xC615;
        private const ushort SprmPIlvl = 0x260A;
        private const ushort SprmPOutLvl = 0x2640;
        private const ushort SprmPIlfo = 0x460B;
        private const ushort SprmPWAlignFont = 0x4439;
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
            out string? warning,
            byte[]? dataStream = null) {
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

                ReadPapxFkp(wordDocumentStream, pageOffset, ranges, dataStream ?? Array.Empty<byte>(), ref warning);
            }

            return ranges
                .OrderBy(range => range.FileOffsetStart)
                .ThenBy(range => range.FileOffsetEnd)
                .ToArray();
        }

        private static void ReadPapxFkp(byte[] wordDocumentStream, int pageOffset, List<LegacyDocParagraphFormatRange> ranges,
            byte[] dataStream, ref string? warning) {
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

                LegacyDocParagraphFormat format;
                try {
                    format = ReadPapx(wordDocumentStream, absolutePapxOffset, pageOffset + OleSectorSize - 1, dataStream);
                } catch (InvalidDataException exception) {
                    warning ??= exception.Message;
                    continue;
                }
                if (format.HasFormatting) {
                    ranges.Add(new LegacyDocParagraphFormatRange(fcStart, fcEnd, format));
                }
            }
        }

        private static LegacyDocParagraphFormat ReadPapx(byte[] bytes, int offset, int pageEnd, byte[] dataStream) {
            if (offset >= pageEnd) {
                return LegacyDocParagraphFormat.Default;
            }

            int cb = bytes[offset];
            int grpprlOffset = offset + 1;
            int grpprlLength = cb * 2 - 1;
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
            byte[] properties = ResolveDataProperties(bytes, grpprlOffset + 2, grpprlLength - 2, dataStream, styleIndex);
            return ReadGrpprl(properties, 0, properties.Length, styleIndex == 0 ? null : styleIndex, requireComplete: true);
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
                case 6:
                    alignment = LegacyDocTabStopAlignment.Number;
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
            return value != 0;
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
