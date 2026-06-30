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
        private const ushort SprmSPgnStart97 = 0x501C;
        private const ushort SprmSDyaHdrTop = 0xB017;
        private const ushort SprmSDyaHdrBottom = 0xB018;
        private const ushort SprmSFTitlePage = 0x300A;
        private const ushort SprmSLBetween = 0x3019;
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
                    case GutterOnRight gutterOnRight:
                        rtlGutter = IsOnOffEnabled(gutterOnRight);
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
                rtlGutter);
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
