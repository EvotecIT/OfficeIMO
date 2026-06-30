using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int DefaultPageWidthTwips = 12240;
        private const int DefaultPageHeightTwips = 15840;
        private const int DefaultPageMarginTwips = 1440;
        private const int DefaultHeaderFooterMarginTwips = 720;
        private const ushort SprmSBkc = 0x3009;
        private const ushort SprmSDyaHdrTop = 0xB017;
        private const ushort SprmSDyaHdrBottom = 0xB018;
        private const ushort SprmSFTitlePage = 0x300A;
        private const ushort SprmSBOrientation = 0x301D;
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
                differentFirstPage);
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

        private static byte[] CreateSepx(LegacyDocSectionFormat sectionFormat) {
            var grpprl = new List<byte>();

            if (sectionFormat.SectionBreakType != null && sectionFormat.SectionBreakType.Value != SectionMarkValues.NextPage) {
                AddSingleByteSprm(grpprl, SprmSBkc, GetSectionBreakTypeOperand(sectionFormat.SectionBreakType.Value));
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
