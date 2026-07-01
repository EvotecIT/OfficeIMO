using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocSectionFormattingReader {
        private const int SedLength = 12;
        private const ushort SprmSBkc = 0x3009;
        private const ushort SprmSCcolumns = 0x500B;
        private const ushort SprmSDxaColumns = 0x900C;
        private const ushort SprmSNfcPgn = 0x300E;
        private const ushort SprmSFPgnRestart = 0x3011;
        private const ushort SprmSLnc = 0x3013;
        private const ushort SprmSFpc = 0x303B;
        private const ushort SprmSRncFtn = 0x303C;
        private const ushort SprmSRncEdn = 0x303E;
        private const ushort SprmSNLnnMod = 0x5015;
        private const ushort SprmSDxaLnn = 0x9016;
        private const ushort SprmSLnnMin = 0x501B;
        private const ushort SprmSNFtn = 0x503F;
        private const ushort SprmSNfcFtnRef = 0x5040;
        private const ushort SprmSNEdn = 0x5041;
        private const ushort SprmSNfcEdnRef = 0x5042;
        private const ushort SprmSPgnStart97 = 0x501C;
        private const ushort SprmSPgnStart = 0x7044;
        private const ushort SprmSDyaHdrTop = 0xB017;
        private const ushort SprmSDyaHdrBottom = 0xB018;
        private const ushort SprmSFTitlePage = 0x300A;
        private const ushort SprmSLBetween = 0x3019;
        private const ushort SprmSVjc = 0x301A;
        private const ushort SprmSBOrientation = 0x301D;
        private const ushort SprmSFRTLGutter = 0x322A;
        private const ushort SprmSBrcTop80 = 0x702B;
        private const ushort SprmSBrcLeft80 = 0x702C;
        private const ushort SprmSBrcBottom80 = 0x702D;
        private const ushort SprmSBrcRight80 = 0x702E;
        private const ushort SprmSPgbProp = 0x522F;
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
            bool rtlGutter = false;
            VerticalJustificationValues? verticalAlignment = null;
            int? lineNumberCountBy = null;
            int? lineNumberDistance = null;
            int? lineNumberStart = null;
            LineNumberRestartValues? lineNumberRestart = null;
            FootnotePositionValues? footnotePosition = null;
            RestartNumberValues? footnoteRestart = null;
            int? footnoteStart = null;
            NumberFormatValues? footnoteNumberFormat = null;
            RestartNumberValues? endnoteRestart = null;
            int? endnoteStart = null;
            NumberFormatValues? endnoteNumberFormat = null;
            SectionMarkValues? sectionBreakType = null;
            LegacyDocParagraphBorder pageTopBorder = default;
            LegacyDocParagraphBorder pageLeftBorder = default;
            LegacyDocParagraphBorder pageBottomBorder = default;
            LegacyDocParagraphBorder pageRightBorder = default;
            LegacyDocPageBorderOptions pageBorderOptions = default;

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

                if (sprm == SprmSFRTLGutter) {
                    if (offset + 3 > end) {
                        break;
                    }

                    rtlGutter = bytes[offset + 2] != 0;
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSVjc) {
                    if (offset + 3 > end) {
                        break;
                    }

                    verticalAlignment = ReadVerticalAlignment(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSLnc) {
                    if (offset + 3 > end) {
                        break;
                    }

                    lineNumberRestart = ReadLineNumberRestart(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSFpc) {
                    if (offset + 3 > end) {
                        break;
                    }

                    footnotePosition = ReadFootnotePosition(bytes[offset + 2]);
                    offset += 3;
                    continue;
                }

                if (sprm == SprmSRncFtn || sprm == SprmSRncEdn) {
                    if (offset + 3 > end) {
                        break;
                    }

                    RestartNumberValues? restart = ReadNoteRestart(bytes[offset + 2]);
                    if (sprm == SprmSRncFtn) {
                        footnoteRestart = restart;
                    } else {
                        endnoteRestart = restart == RestartNumberValues.EachPage ? null : restart;
                    }

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

                if (sprm == SprmSBrcTop80
                    || sprm == SprmSBrcLeft80
                    || sprm == SprmSBrcBottom80
                    || sprm == SprmSBrcRight80) {
                    if (offset + 6 > end) {
                        break;
                    }

                    LegacyDocParagraphBorder border = ReadBrc80Border(bytes, offset + 2);
                    switch (sprm) {
                        case SprmSBrcTop80:
                            pageTopBorder = border;
                            break;
                        case SprmSBrcLeft80:
                            pageLeftBorder = border;
                            break;
                        case SprmSBrcBottom80:
                            pageBottomBorder = border;
                            break;
                        case SprmSBrcRight80:
                            pageRightBorder = border;
                            break;
                    }

                    offset += 6;
                    continue;
                }

                if (sprm == SprmSPgbProp) {
                    if (offset + 4 > end) {
                        break;
                    }

                    pageBorderOptions = ReadPageBorderOptions(bytes[offset + 2]);
                    offset += 4;
                    continue;
                }

                if (sprm == SprmSDyaHdrTop
                    || sprm == SprmSDyaHdrBottom
                    || sprm == SprmSCcolumns
                    || sprm == SprmSDxaColumns
                    || sprm == SprmSNLnnMod
                    || sprm == SprmSDxaLnn
                    || sprm == SprmSLnnMin
                    || sprm == SprmSNFtn
                    || sprm == SprmSNfcFtnRef
                    || sprm == SprmSNEdn
                    || sprm == SprmSNfcEdnRef
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
                        case SprmSNLnnMod:
                            lineNumberCountBy = value == 0 ? null : value;
                            break;
                        case SprmSDxaLnn:
                            lineNumberDistance = value;
                            break;
                        case SprmSLnnMin:
                            lineNumberStart = value + 1;
                            break;
                        case SprmSNFtn:
                            footnoteStart = value == 0 ? null : value;
                            break;
                        case SprmSNfcFtnRef:
                            footnoteNumberFormat = ReadPageNumberFormat(bytes[offset + 2]);
                            break;
                        case SprmSNEdn:
                            endnoteStart = value == 0 ? null : value;
                            break;
                        case SprmSNfcEdnRef:
                            endnoteNumberFormat = ReadPageNumberFormat(bytes[offset + 2]);
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

            return new LegacyDocSectionFormat(
                sectionBreakType,
                pageWidth,
                pageHeight,
                orientation,
                marginTop,
                marginRight,
                marginBottom,
                marginLeft,
                headerDistance,
                footerDistance,
                gutter,
                differentFirstPage,
                columnCount,
                columnSpacing,
                hasColumnSeparator,
                restartPageNumbering ? pageNumberStart ?? 0 : null,
                pageNumberFormat,
                rtlGutter,
                verticalAlignment,
                lineNumberCountBy,
                lineNumberDistance,
                lineNumberStart,
                lineNumberRestart,
                footnotePosition,
                footnoteRestart,
                footnoteStart,
                footnoteNumberFormat,
                null,
                endnoteRestart,
                endnoteStart,
                endnoteNumberFormat,
                new LegacyDocParagraphBorders(pageTopBorder, pageLeftBorder, pageBottomBorder, pageRightBorder, default, pageBorderOptions));
        }

        private static FootnotePositionValues? ReadFootnotePosition(byte value) {
            switch (value) {
                case 1:
                    return FootnotePositionValues.PageBottom;
                case 2:
                    return FootnotePositionValues.BeneathText;
                default:
                    return null;
            }
        }

        private static RestartNumberValues? ReadNoteRestart(byte value) {
            switch (value) {
                case 0:
                    return RestartNumberValues.Continuous;
                case 1:
                    return RestartNumberValues.EachSection;
                case 2:
                    return RestartNumberValues.EachPage;
                default:
                    return null;
            }
        }

        private static LineNumberRestartValues? ReadLineNumberRestart(byte value) {
            switch (value) {
                case 0:
                    return LineNumberRestartValues.NewPage;
                case 1:
                    return LineNumberRestartValues.NewSection;
                case 2:
                    return LineNumberRestartValues.Continuous;
                default:
                    return null;
            }
        }

        private static VerticalJustificationValues? ReadVerticalAlignment(byte value) {
            switch (value) {
                case 0:
                    return VerticalJustificationValues.Top;
                case 1:
                    return VerticalJustificationValues.Center;
                case 2:
                    return VerticalJustificationValues.Both;
                case 3:
                    return VerticalJustificationValues.Bottom;
                default:
                    return null;
            }
        }

        private static NumberFormatValues? ReadPageNumberFormat(byte value) {
            return LegacyDocNumberFormatMapper.FromNfc(value);
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

        private static LegacyDocParagraphBorder ReadBrc80Border(byte[] bytes, int offset) {
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
            LegacyDocParagraphBorderStyle style = MapBrc80BorderStyle(borderType);
            if (style == LegacyDocParagraphBorderStyle.None || sizeEighthPoints == 0) {
                return default;
            }

            string? colorHex = LegacyDocColorPalette.GetHexForIco(colorIndex);
            return new LegacyDocParagraphBorder(style, colorHex, sizeEighthPoints, spacePoints);
        }

        private static LegacyDocPageBorderOptions ReadPageBorderOptions(byte operand) {
            return new LegacyDocPageBorderOptions(
                ReadPageBorderDisplay(operand & 0x07),
                ReadPageBorderOffsetFrom((operand >> 5) & 0x07),
                ReadPageBorderZOrder((operand >> 3) & 0x03));
        }

        private static LegacyDocPageBorderDisplay ReadPageBorderDisplay(int value) {
            switch (value) {
                case 0:
                    return LegacyDocPageBorderDisplay.AllPages;
                case 1:
                    return LegacyDocPageBorderDisplay.FirstPage;
                case 2:
                    return LegacyDocPageBorderDisplay.NotFirstPage;
                default:
                    return LegacyDocPageBorderDisplay.AllPages;
            }
        }

        private static LegacyDocPageBorderOffsetFrom ReadPageBorderOffsetFrom(int value) {
            switch (value) {
                case 0:
                    return LegacyDocPageBorderOffsetFrom.Text;
                case 1:
                    return LegacyDocPageBorderOffsetFrom.Page;
                default:
                    return LegacyDocPageBorderOffsetFrom.Text;
            }
        }

        private static LegacyDocPageBorderZOrder ReadPageBorderZOrder(int value) {
            switch (value) {
                case 0:
                    return LegacyDocPageBorderZOrder.Front;
                case 1:
                    return LegacyDocPageBorderZOrder.Back;
                default:
                    return LegacyDocPageBorderZOrder.Front;
            }
        }

        private static LegacyDocParagraphBorderStyle MapBrc80BorderStyle(byte borderType) {
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
