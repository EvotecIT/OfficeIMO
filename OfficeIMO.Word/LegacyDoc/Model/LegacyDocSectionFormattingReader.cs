using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocSectionFormattingReader {
        private const int SedLength = 12;
        private const ushort SprmSBkc = 0x3009;
        private const ushort SprmSCcolumns = 0x500B;
        private const ushort SprmSDxaColumns = 0x900C;
        private const ushort SprmSNfcPgn = 0x300E;
        private const ushort SprmSFPgnRestart = 0x3011;
        private const ushort SprmSPgnStart97 = 0x501C;
        private const ushort SprmSPgnStart = 0x7044;
        private const ushort SprmSDyaHdrTop = 0xB017;
        private const ushort SprmSDyaHdrBottom = 0xB018;
        private const ushort SprmSFTitlePage = 0x300A;
        private const ushort SprmSLBetween = 0x3019;
        private const ushort SprmSBOrientation = 0x301D;
        private const ushort SprmSXaPage = 0xB01F;
        private const ushort SprmSYaPage = 0xB020;
        private const ushort SprmSDxaLeft = 0xB021;
        private const ushort SprmSDxaRight = 0xB022;
        private const ushort SprmSDyaTop = 0x9023;
        private const ushort SprmSDyaBottom = 0x9024;
        private const ushort SprmSDzaGutter = 0xB025;

        internal static LegacyDocSectionFormat ReadSectionFormatting(byte[] wordDocumentStream, byte[] tableStream, LegacyDocFib fib, out string? warning) {
            IReadOnlyList<LegacyDocSection> sections = ReadSections(wordDocumentStream, tableStream, fib, out warning);
            return sections.Count == 0
                ? LegacyDocSectionFormat.Default
                : sections[0].Format;
        }

        internal static IReadOnlyList<LegacyDocSection> ReadSections(byte[] wordDocumentStream, byte[] tableStream, LegacyDocFib fib, out string? warning) {
            warning = null;
            if (fib.LcbPlcfSed == 0) {
                return new[] { new LegacyDocSection(0, Math.Max(0, fib.CcpText), LegacyDocSectionFormat.Default) };
            }

            if (fib.FcPlcfSed < 0
                || fib.LcbPlcfSed < 0
                || fib.FcPlcfSed + fib.LcbPlcfSed > tableStream.Length
                || fib.LcbPlcfSed < 20
                || (fib.LcbPlcfSed - 4) % 16 != 0) {
                warning = "The FIB points outside the selected table stream for the section descriptor PLC.";
                return new[] { new LegacyDocSection(0, Math.Max(0, fib.CcpText), LegacyDocSectionFormat.Default) };
            }

            int sectionCount = (fib.LcbPlcfSed - 4) / 16;
            if (sectionCount <= 0) {
                return new[] { new LegacyDocSection(0, Math.Max(0, fib.CcpText), LegacyDocSectionFormat.Default) };
            }

            int sectionDescriptorOffset = fib.FcPlcfSed + ((sectionCount + 1) * 4);
            if (sectionDescriptorOffset + (sectionCount * SedLength) > fib.FcPlcfSed + fib.LcbPlcfSed) {
                warning = "The section descriptor PLC does not contain a complete section descriptor.";
                return new[] { new LegacyDocSection(0, Math.Max(0, fib.CcpText), LegacyDocSectionFormat.Default) };
            }

            var sections = new List<LegacyDocSection>(sectionCount);
            for (int sectionIndex = 0; sectionIndex < sectionCount; sectionIndex++) {
                int startCharacter = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcfSed + (sectionIndex * 4));
                int endCharacter = LegacyDocFib.ReadInt32(tableStream, fib.FcPlcfSed + ((sectionIndex + 1) * 4));
                if (startCharacter < 0 || endCharacter < startCharacter) {
                    warning = "The section descriptor PLC contains an invalid character range.";
                    return new[] { new LegacyDocSection(0, Math.Max(0, fib.CcpText), LegacyDocSectionFormat.Default) };
                }

                int sedOffset = sectionDescriptorOffset + (sectionIndex * SedLength);
                LegacyDocSectionFormat format = ReadSectionFormat(wordDocumentStream, tableStream, sedOffset, out string? sectionWarning);
                if (sectionWarning != null) {
                    warning = sectionWarning;
                }

                sections.Add(new LegacyDocSection(startCharacter, endCharacter, format));
            }

            return sections;
        }

        private static LegacyDocSectionFormat ReadSectionFormat(byte[] wordDocumentStream, byte[] tableStream, int sedOffset, out string? warning) {
            warning = null;
            int fcSepx = LegacyDocFib.ReadInt32(tableStream, sedOffset + 2);
            if (fcSepx <= 0) {
                return LegacyDocSectionFormat.Default;
            }

            if (fcSepx + 2 > wordDocumentStream.Length) {
                warning = "The section descriptor points outside the WordDocument stream.";
                return LegacyDocSectionFormat.Default;
            }

            int cb = LegacyDocFib.ReadUInt16(wordDocumentStream, fcSepx);
            if (cb < 0 || fcSepx + 2 + cb > wordDocumentStream.Length) {
                warning = "The section property block points outside the WordDocument stream.";
                return LegacyDocSectionFormat.Default;
            }

            return ReadSepxGrpprl(wordDocumentStream, fcSepx + 2, cb);
        }

        private static LegacyDocSectionFormat ReadSepxGrpprl(byte[] bytes, int offset, int count) {
            int end = offset + count;
            int? pageWidth = null;
            int? pageHeight = null;
            PageOrientationValues? orientation = null;
            int? marginTop = null;
            int? marginRight = null;
            int? marginBottom = null;
            int? marginLeft = null;
            int? headerDistance = null;
            int? footerDistance = null;
            int? gutter = null;
            bool differentFirstPage = false;
            int? columnCount = null;
            int? columnSpacing = null;
            bool hasColumnSeparator = false;
            bool restartPageNumbering = false;
            int? pageNumberStart = null;
            NumberFormatValues? pageNumberFormat = null;
            SectionMarkValues? sectionBreakType = null;

            while (offset + 2 <= end) {
                ushort sprm = LegacyDocFib.ReadUInt16(bytes, offset);
                if (sprm == SprmSBkc) {
                    if (offset + 3 > end) {
                        break;
                    }

                    sectionBreakType = ReadSectionBreakType(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSBOrientation) {
                    if (offset + 3 > end) {
                        break;
                    }

                    orientation = bytes[offset + 2] == 2 ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSFTitlePage) {
                    if (offset + 3 > end) {
                        break;
                    }

                    differentFirstPage = bytes[offset + 2] != 0;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSLBetween) {
                    if (offset + 3 > end) {
                        break;
                    }

                    hasColumnSeparator = bytes[offset + 2] != 0;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSNfcPgn) {
                    if (offset + 3 > end) {
                        break;
                    }

                    pageNumberFormat = ReadPageNumberFormat(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSFPgnRestart) {
                    if (offset + 3 > end) {
                        break;
                    }

                    restartPageNumbering = bytes[offset + 2] != 0;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSDyaHdrTop
                    || sprm == SprmSDyaHdrBottom
                    || sprm == SprmSCcolumns
                    || sprm == SprmSDxaColumns
                    || sprm == SprmSPgnStart97
                    || sprm == SprmSXaPage
                    || sprm == SprmSYaPage
                    || sprm == SprmSDxaLeft
                    || sprm == SprmSDxaRight
                    || sprm == SprmSDyaTop
                    || sprm == SprmSDyaBottom
                    || sprm == SprmSDzaGutter) {
                    if (offset + 4 > end) {
                        break;
                    }

                    int value = ReadUInt16AsInt(bytes, offset + 2);
                    switch (sprm) {
                        case SprmSDyaHdrTop:
                            headerDistance = value;
                            break;
                        case SprmSDyaHdrBottom:
                            footerDistance = value;
                            break;
                        case SprmSCcolumns:
                            columnCount = value + 1;
                            break;
                        case SprmSDxaColumns:
                            columnSpacing = value;
                            break;
                        case SprmSPgnStart97:
                            pageNumberStart = value;
                            break;
                        case SprmSXaPage:
                            pageWidth = value;
                            break;
                        case SprmSYaPage:
                            pageHeight = value;
                            break;
                        case SprmSDxaLeft:
                            marginLeft = value;
                            break;
                        case SprmSDxaRight:
                            marginRight = value;
                            break;
                        case SprmSDyaTop:
                            marginTop = value;
                            break;
                        case SprmSDyaBottom:
                            marginBottom = value;
                            break;
                        case SprmSDzaGutter:
                            gutter = value;
                            break;
                    }

                    offset += 4;
                    continue;
                }

                if (sprm == SprmSPgnStart) {
                    if (offset + 6 > end) {
                        break;
                    }

                    int value = LegacyDocFib.ReadInt32(bytes, offset + 2);
                    if (value >= 0) {
                        pageNumberStart = value;
                    }

                    offset += 6;
                    continue;
                }

                if (!TryGetSprmOperandLength(bytes, offset, end, out int operandLength)) {
                    break;
                }

                offset += 2 + operandLength;
            }

            if (orientation == null && pageWidth != null && pageHeight != null && pageWidth > pageHeight) {
                orientation = PageOrientationValues.Landscape;
            }

            return new LegacyDocSectionFormat(sectionBreakType, pageWidth, pageHeight, orientation, marginTop, marginRight, marginBottom, marginLeft, headerDistance, footerDistance, gutter, differentFirstPage, columnCount, columnSpacing, hasColumnSeparator, restartPageNumbering ? pageNumberStart ?? 0 : null, pageNumberFormat);
        }

        private static NumberFormatValues? ReadPageNumberFormat(byte value) {
            switch (value) {
                case 0:
                    return NumberFormatValues.Decimal;
                case 1:
                    return NumberFormatValues.UpperRoman;
                case 2:
                    return NumberFormatValues.LowerRoman;
                case 3:
                    return NumberFormatValues.UpperLetter;
                case 4:
                    return NumberFormatValues.LowerLetter;
                default:
                    return null;
            }
        }

        private static SectionMarkValues? ReadSectionBreakType(byte value) {
            switch (value) {
                case 0:
                    return SectionMarkValues.Continuous;
                case 1:
                    return SectionMarkValues.NextColumn;
                case 2:
                    return SectionMarkValues.NextPage;
                case 3:
                    return SectionMarkValues.EvenPage;
                case 4:
                    return SectionMarkValues.OddPage;
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

        private static int ReadUInt16AsInt(byte[] bytes, int offset) {
            return LegacyDocFib.ReadUInt16(bytes, offset);
        }
    }
}
