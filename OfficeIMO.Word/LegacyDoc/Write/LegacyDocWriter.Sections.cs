using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int DefaultPageWidthTwips = 12240;
        private const int DefaultPageHeightTwips = 15840;
        private const int DefaultPageMarginTwips = 1440;
        private const int DefaultHeaderFooterMarginTwips = 720;
        private const int DefaultColumnSpaceTwips = 720;
        private const int MaxLegacySectionColumns = 45;
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
        private const ushort SprmSDyaHdrTop = 0xB017;
        private const ushort SprmSDyaHdrBottom = 0xB018;
        private const ushort SprmSFTitlePage = 0x300A;
        private const ushort SprmSLBetween = 0x3019;
        private const ushort SprmSVjc = 0x301A;
        private const ushort SprmSBOrientation = 0x301D;
        private const ushort SprmSFRTLGutter = 0x322A;
        private const ushort SprmSXaPage = 0xB01F;
        private const ushort SprmSYaPage = 0xB020;
        private const ushort SprmSDxaLeft = 0xB021;
        private const ushort SprmSDxaRight = 0xB022;
        private const ushort SprmSDyaTop = 0x9023;
        private const ushort SprmSDyaBottom = 0x9024;
        private const ushort SprmSDzaGutter = 0xB025;

        private static LegacyDocSectionFormat ReadSupportedSectionProperties(SectionProperties sectionProperties) {
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

            foreach (OpenXmlElement property in sectionProperties.ChildElements) {
                switch (property) {
                    case HeaderReference:
                    case FooterReference:
                        break;
                    case TitlePage titlePage:
                        differentFirstPage = IsOnOffEnabled(titlePage);
                        break;
                    case PageSize pageSize:
                        pageWidth = ReadTwipValue(pageSize.Width, DefaultPageWidthTwips, "section page width");
                        pageHeight = ReadTwipValue(pageSize.Height, DefaultPageHeightTwips, "section page height");
                        orientation = pageSize.Orient?.Value;
                        if (orientation == null && pageWidth > pageHeight) {
                            orientation = PageOrientationValues.Landscape;
                        }

                        break;
                    case PageMargin pageMargin:
                        marginTop = ReadTwipValue(pageMargin.Top, DefaultPageMarginTwips, "section top margin");
                        marginRight = ReadTwipValue(pageMargin.Right, DefaultPageMarginTwips, "section right margin");
                        marginBottom = ReadTwipValue(pageMargin.Bottom, DefaultPageMarginTwips, "section bottom margin");
                        marginLeft = ReadTwipValue(pageMargin.Left, DefaultPageMarginTwips, "section left margin");
                        headerDistance = ReadTwipValue(pageMargin.Header, DefaultHeaderFooterMarginTwips, "section header distance");
                        footerDistance = ReadTwipValue(pageMargin.Footer, DefaultHeaderFooterMarginTwips, "section footer distance");
                        gutter = ReadTwipValue(pageMargin.Gutter, 0, "section gutter");
                        break;
                    case SectionType sectionType:
                        if (sectionType.Val != null) {
                            sectionBreakType = sectionType.Val.Value;
                        }

                        break;
                    case Columns columns:
                        ReadSupportedColumns(columns, out columnCount, out columnSpacing, out hasColumnSeparator);
                        break;
                    case PageNumberType pageNumberType:
                        pageNumberStart = ReadPageNumberStart(pageNumberType.Start);
                        pageNumberFormat = ReadPageNumberFormat(pageNumberType.Format);
                        break;
                    case LineNumberType lineNumberType:
                        ReadSupportedLineNumbering(lineNumberType, out lineNumberCountBy, out lineNumberDistance, out lineNumberStart, out lineNumberRestart);
                        break;
                    case FootnoteProperties footnoteProperties:
                        ReadSupportedFootnoteProperties(footnoteProperties, out footnotePosition, out footnoteRestart, out footnoteStart, out footnoteNumberFormat);
                        break;
                    case EndnoteProperties endnoteProperties:
                        ReadSupportedEndnoteProperties(endnoteProperties, out endnoteRestart, out endnoteStart, out endnoteNumberFormat);
                        break;
                    case GutterOnRight gutterOnRight:
                        rtlGutter = IsOnOffEnabled(gutterOnRight);
                        break;
                    case VerticalTextAlignmentOnPage verticalTextAlignment:
                        verticalAlignment = ReadVerticalAlignment(verticalTextAlignment.Val);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports simple section page setup only. Unsupported section property: {property.LocalName}.");
                }
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
                pageNumberStart,
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
                endnoteRestart,
                endnoteStart,
                endnoteNumberFormat);
        }

        private static void ReadSupportedColumns(Columns columns, out int? columnCount, out int? columnSpacing, out bool hasColumnSeparator) {
            if ((columns.EqualWidth != null && !columns.EqualWidth.Value) || columns.Elements<Column>().Any()) {
                throw new NotSupportedException("Native DOC saving supports equal-width section columns only.");
            }

            columnCount = ReadColumnCount(columns.ColumnCount);
            columnSpacing = ReadColumnSpacing(columns.Space, columnCount != null ? DefaultColumnSpaceTwips : null);
            hasColumnSeparator = columns.Separator?.Value ?? false;
        }

        private static bool IsOnOffEnabled(OnOffType element) {
            return element.Val == null || element.Val.Value;
        }

        private static void ReadSupportedLineNumbering(
            LineNumberType lineNumberType,
            out int? countBy,
            out int? distance,
            out int? start,
            out LineNumberRestartValues? restart) {
            countBy = ReadLineNumberCountBy(lineNumberType.CountBy);
            distance = ReadLineNumberDistance(lineNumberType.Distance);
            start = ReadLineNumberStart(lineNumberType.Start);
            restart = ReadLineNumberRestart(lineNumberType.Restart);
        }

        private static void ReadSupportedFootnoteProperties(
            FootnoteProperties footnoteProperties,
            out FootnotePositionValues? position,
            out RestartNumberValues? restart,
            out int? start,
            out NumberFormatValues? numberFormat) {
            position = null;
            restart = null;
            start = null;
            numberFormat = null;

            foreach (OpenXmlElement property in footnoteProperties.ChildElements) {
                switch (property) {
                    case FootnotePosition footnotePosition:
                        position = ReadFootnotePosition(footnotePosition.Val);
                        break;
                    case NumberingRestart numberingRestart:
                        restart = ReadFootnoteRestart(numberingRestart.Val);
                        break;
                    case NumberingStart numberingStart:
                        start = ReadNoteNumberStart(numberingStart.Val, "section footnote start");
                        break;
                    case NumberingFormat numberingFormat:
                        numberFormat = ReadPageNumberFormat(numberingFormat.Val);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving does not support section footnote property '{property.LocalName}'.");
                }
            }
        }

        private static void ReadSupportedEndnoteProperties(
            EndnoteProperties endnoteProperties,
            out RestartNumberValues? restart,
            out int? start,
            out NumberFormatValues? numberFormat) {
            restart = null;
            start = null;
            numberFormat = null;

            foreach (OpenXmlElement property in endnoteProperties.ChildElements) {
                switch (property) {
                    case EndnotePosition:
                        throw new NotSupportedException("Native DOC saving does not support section endnote placement yet.");
                    case NumberingRestart numberingRestart:
                        restart = ReadEndnoteRestart(numberingRestart.Val);
                        break;
                    case NumberingStart numberingStart:
                        start = ReadNoteNumberStart(numberingStart.Val, "section endnote start");
                        break;
                    case NumberingFormat numberingFormat:
                        numberFormat = ReadPageNumberFormat(numberingFormat.Val);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving does not support section endnote property '{property.LocalName}'.");
                }
            }
        }

        private static int? ReadTwipValue(OpenXmlSimpleType? value, int defaultValue, string description) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > ushort.MaxValue) {
                throw new NotSupportedException($"Native DOC saving supports {description} only within the Word 97-2003 unsigned twip range.");
            }

            return actual == defaultValue ? null : actual;
        }

        private static int? ReadColumnCount(OpenXmlSimpleType? value) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 1
                || actual > MaxLegacySectionColumns) {
                throw new NotSupportedException($"Native DOC saving supports section column counts from 1 through {MaxLegacySectionColumns}.");
            }

            return actual;
        }

        private static int? ReadColumnSpacing(OpenXmlSimpleType? value, int? defaultValue) {
            if (value == null) {
                return defaultValue;
            }

            if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section column spacing only within the Word 97-2003 unsigned twip range.");
            }

            return actual;
        }

        private static int? ReadPageNumberStart(OpenXmlSimpleType? value) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section page number starts only within the Word 97-2003 unsigned range.");
            }

            return actual;
        }

        private static NumberFormatValues? ReadPageNumberFormat(EnumValue<NumberFormatValues>? value) {
            if (value == null) {
                return null;
            }

            return GetPageNumberFormatOperand(value.Value) != null
                ? value.Value
                : throw new NotSupportedException($"Native DOC saving does not support section page number format '{value.Value}'.");
        }

        private static FootnotePositionValues? ReadFootnotePosition(EnumValue<FootnotePositionValues>? value) {
            if (value == null) {
                return null;
            }

            return GetFootnotePositionOperand(value.Value) != null
                ? value.Value
                : throw new NotSupportedException($"Native DOC saving does not support section footnote placement '{value.Value}'.");
        }

        private static RestartNumberValues? ReadFootnoteRestart(EnumValue<RestartNumberValues>? value) {
            if (value == null) {
                return null;
            }

            return GetNoteRestartOperand(value.Value) != null
                ? value.Value
                : throw new NotSupportedException($"Native DOC saving does not support section footnote numbering restart '{value.Value}'.");
        }

        private static RestartNumberValues? ReadEndnoteRestart(EnumValue<RestartNumberValues>? value) {
            if (value == null) {
                return null;
            }

            if (value.Value == RestartNumberValues.EachPage) {
                throw new NotSupportedException("Native DOC saving does not support section endnote numbering restart for each page.");
            }

            return GetNoteRestartOperand(value.Value) != null
                ? value.Value
                : throw new NotSupportedException($"Native DOC saving does not support section endnote numbering restart '{value.Value}'.");
        }

        private static int? ReadNoteNumberStart(OpenXmlSimpleType? value, string description) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 1
                || actual > 16383) {
                throw new NotSupportedException($"Native DOC saving supports {description} only from 1 through 16383.");
            }

            return actual;
        }

        private static int? ReadLineNumberCountBy(OpenXmlSimpleType? value) {
            if (value == null) {
                return 1;
            }

            if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > 100) {
                throw new NotSupportedException("Native DOC saving supports section line number intervals from 0 through 100.");
            }

            return actual;
        }

        private static int? ReadLineNumberDistance(OpenXmlSimpleType? value) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section line number distance only within the Word 97-2003 unsigned twip range.");
            }

            return actual;
        }

        private static int? ReadLineNumberStart(OpenXmlSimpleType? value) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 1
                || actual > 32767) {
                throw new NotSupportedException("Native DOC saving supports section line number starts from 1 through 32767.");
            }

            return actual;
        }

        private static LineNumberRestartValues? ReadLineNumberRestart(EnumValue<LineNumberRestartValues>? value) {
            if (value == null) {
                return null;
            }

            return GetLineNumberRestartOperand(value.Value) != null
                ? value.Value
                : throw new NotSupportedException($"Native DOC saving does not support section line number restart mode '{value.Value}'.");
        }

        private static VerticalJustificationValues? ReadVerticalAlignment(EnumValue<VerticalJustificationValues>? value) {
            if (value == null || value.Value == VerticalJustificationValues.Top) {
                return null;
            }

            return GetVerticalAlignmentOperand(value.Value) != null
                ? value.Value
                : throw new NotSupportedException($"Native DOC saving does not support section vertical alignment '{value.Value}'.");
        }

        private static byte[] CreateSepx(LegacyDocSectionFormat sectionFormat) {
            var grpprl = new List<byte>();

            if (sectionFormat.SectionBreakType != null && sectionFormat.SectionBreakType.Value != SectionMarkValues.NextPage) {
                AddSingleByteSprm(grpprl, SprmSBkc, GetSectionBreakTypeOperand(sectionFormat.SectionBreakType.Value));
            }

            if (sectionFormat.ColumnCount != null) {
                AddUInt16Sprm(grpprl, SprmSCcolumns, sectionFormat.ColumnCount.Value - 1);
            }

            if (sectionFormat.ColumnSpacingTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDxaColumns, sectionFormat.ColumnSpacingTwips.Value);
            }

            if (sectionFormat.HasColumnSeparator) {
                AddSingleByteSprm(grpprl, SprmSLBetween, 1);
            }

            if (sectionFormat.PageNumberFormat != null) {
                AddSingleByteSprm(grpprl, SprmSNfcPgn, GetPageNumberFormatOperand(sectionFormat.PageNumberFormat.Value)!.Value);
            }

            if (sectionFormat.PageNumberStart != null) {
                AddSingleByteSprm(grpprl, SprmSFPgnRestart, 1);
                AddUInt16Sprm(grpprl, SprmSPgnStart97, sectionFormat.PageNumberStart.Value);
            }

            if (sectionFormat.FootnotePosition != null) {
                AddSingleByteSprm(grpprl, SprmSFpc, GetFootnotePositionOperand(sectionFormat.FootnotePosition.Value)!.Value);
            }

            if (sectionFormat.FootnoteRestart != null) {
                AddSingleByteSprm(grpprl, SprmSRncFtn, GetNoteRestartOperand(sectionFormat.FootnoteRestart.Value)!.Value);
            }

            if (sectionFormat.EndnoteRestart != null) {
                AddSingleByteSprm(grpprl, SprmSRncEdn, GetNoteRestartOperand(sectionFormat.EndnoteRestart.Value)!.Value);
            }

            if (sectionFormat.FootnoteStart != null) {
                AddUInt16Sprm(grpprl, SprmSNFtn, sectionFormat.FootnoteStart.Value);
            }

            if (sectionFormat.FootnoteNumberFormat != null) {
                AddUInt16Sprm(grpprl, SprmSNfcFtnRef, GetPageNumberFormatOperand(sectionFormat.FootnoteNumberFormat.Value)!.Value);
            }

            if (sectionFormat.EndnoteStart != null) {
                AddUInt16Sprm(grpprl, SprmSNEdn, sectionFormat.EndnoteStart.Value);
            }

            if (sectionFormat.EndnoteNumberFormat != null) {
                AddUInt16Sprm(grpprl, SprmSNfcEdnRef, GetPageNumberFormatOperand(sectionFormat.EndnoteNumberFormat.Value)!.Value);
            }

            if (sectionFormat.LineNumberRestart != null) {
                AddSingleByteSprm(grpprl, SprmSLnc, GetLineNumberRestartOperand(sectionFormat.LineNumberRestart.Value)!.Value);
            }

            if (sectionFormat.LineNumberCountBy != null) {
                AddUInt16Sprm(grpprl, SprmSNLnnMod, sectionFormat.LineNumberCountBy.Value);
            }

            if (sectionFormat.LineNumberDistanceTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDxaLnn, sectionFormat.LineNumberDistanceTwips.Value);
            }

            if (sectionFormat.LineNumberStart != null) {
                AddUInt16Sprm(grpprl, SprmSLnnMin, sectionFormat.LineNumberStart.Value - 1);
            }

            if (sectionFormat.HeaderDistanceTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDyaHdrTop, sectionFormat.HeaderDistanceTwips.Value);
            }

            if (sectionFormat.FooterDistanceTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDyaHdrBottom, sectionFormat.FooterDistanceTwips.Value);
            }

            if (sectionFormat.DifferentFirstPage) {
                AddSingleByteSprm(grpprl, SprmSFTitlePage, 1);
            }

            if (sectionFormat.Orientation == PageOrientationValues.Landscape) {
                AddSingleByteSprm(grpprl, SprmSBOrientation, 2);
            }

            if (sectionFormat.PageWidthTwips != null) {
                AddUInt16Sprm(grpprl, SprmSXaPage, sectionFormat.PageWidthTwips.Value);
            }

            if (sectionFormat.PageHeightTwips != null) {
                AddUInt16Sprm(grpprl, SprmSYaPage, sectionFormat.PageHeightTwips.Value);
            }

            if (sectionFormat.MarginLeftTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDxaLeft, sectionFormat.MarginLeftTwips.Value);
            }

            if (sectionFormat.MarginRightTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDxaRight, sectionFormat.MarginRightTwips.Value);
            }

            if (sectionFormat.MarginTopTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDyaTop, sectionFormat.MarginTopTwips.Value);
            }

            if (sectionFormat.MarginBottomTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDyaBottom, sectionFormat.MarginBottomTwips.Value);
            }

            if (sectionFormat.GutterTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDzaGutter, sectionFormat.GutterTwips.Value);
            }

            if (sectionFormat.RtlGutter) {
                AddSingleByteSprm(grpprl, SprmSFRTLGutter, 1);
            }

            if (sectionFormat.VerticalAlignment != null) {
                AddSingleByteSprm(grpprl, SprmSVjc, GetVerticalAlignmentOperand(sectionFormat.VerticalAlignment.Value)!.Value);
            }

            if (grpprl.Count > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving cannot write section page setup because the SEPX record is too large.");
            }

            var sepx = new byte[2 + grpprl.Count];
            sepx[0] = (byte)(grpprl.Count & 0xFF);
            sepx[1] = (byte)(grpprl.Count >> 8);
            grpprl.CopyTo(sepx, 2);
            return sepx;
        }

        private static byte GetSectionBreakTypeOperand(SectionMarkValues sectionBreakType) {
            if (sectionBreakType == SectionMarkValues.Continuous) {
                return 0;
            }

            if (sectionBreakType == SectionMarkValues.NextColumn) {
                return 1;
            }

            if (sectionBreakType == SectionMarkValues.NextPage) {
                return 2;
            }

            if (sectionBreakType == SectionMarkValues.EvenPage) {
                return 3;
            }

            if (sectionBreakType == SectionMarkValues.OddPage) {
                return 4;
            }

            throw new NotSupportedException($"Native DOC saving does not support section break type '{sectionBreakType}'.");
        }

        private static byte? GetPageNumberFormatOperand(NumberFormatValues format) {
            if (format == NumberFormatValues.Decimal) {
                return 0;
            }

            if (format == NumberFormatValues.UpperRoman) {
                return 1;
            }

            if (format == NumberFormatValues.LowerRoman) {
                return 2;
            }

            if (format == NumberFormatValues.UpperLetter) {
                return 3;
            }

            if (format == NumberFormatValues.LowerLetter) {
                return 4;
            }

            return null;
        }

        private static byte? GetLineNumberRestartOperand(LineNumberRestartValues restart) {
            if (restart == LineNumberRestartValues.NewPage) {
                return 0;
            }

            if (restart == LineNumberRestartValues.NewSection) {
                return 1;
            }

            if (restart == LineNumberRestartValues.Continuous) {
                return 2;
            }

            return null;
        }

        private static byte? GetFootnotePositionOperand(FootnotePositionValues position) {
            if (position == FootnotePositionValues.PageBottom) {
                return 1;
            }

            if (position == FootnotePositionValues.BeneathText) {
                return 2;
            }

            return null;
        }

        private static byte? GetNoteRestartOperand(RestartNumberValues restart) {
            if (restart == RestartNumberValues.Continuous) {
                return 0;
            }

            if (restart == RestartNumberValues.EachSection) {
                return 1;
            }

            if (restart == RestartNumberValues.EachPage) {
                return 2;
            }

            return null;
        }

        private static byte? GetVerticalAlignmentOperand(VerticalJustificationValues alignment) {
            if (alignment == VerticalJustificationValues.Top) {
                return 0;
            }

            if (alignment == VerticalJustificationValues.Center) {
                return 1;
            }

            if (alignment == VerticalJustificationValues.Both) {
                return 2;
            }

            if (alignment == VerticalJustificationValues.Bottom) {
                return 3;
            }

            return null;
        }

        private static void WritePlcfSed(byte[] table, int offset, IReadOnlyList<LegacyDocWritableSectionRecord> sectionRecords) {
            WriteInt32(table, offset, 0);
            for (int index = 0; index < sectionRecords.Count; index++) {
                WriteInt32(table, offset + ((index + 1) * 4), sectionRecords[index].EndCharacter);
            }

            int sedOffset = offset + ((sectionRecords.Count + 1) * 4);
            for (int index = 0; index < sectionRecords.Count; index++) {
                int recordOffset = sedOffset + (index * SedLength);
                WriteUInt16(table, recordOffset, 0);
                WriteInt32(table, recordOffset + 2, sectionRecords[index].SepxOffset);
                WriteUInt16(table, recordOffset + 6, 0);
                WriteUInt16(table, recordOffset + 8, 0);
                WriteUInt16(table, recordOffset + 10, 0);
            }
        }

        private static void AddUInt16Sprm(List<byte> grpprl, ushort sprm, int operand) {
            if (operand < 0 || operand > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section page setup values only within the Word 97-2003 unsigned twip range.");
            }

            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)(operand & 0xFF));
            grpprl.Add((byte)(operand >> 8));
        }
    }
}
