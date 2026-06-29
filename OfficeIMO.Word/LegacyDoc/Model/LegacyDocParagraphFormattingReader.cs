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
        private const ushort SprmPFWidowControl = 0x2431;
        private const ushort SprmPChgTabsPapx = 0xC60D;

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
            bool? isInTable = null;
            bool? isTableTerminatingParagraph = null;
            var tabStops = new List<LegacyDocTabStop>();
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
                isInTable,
                isTableTerminatingParagraph,
                tabStops);
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
