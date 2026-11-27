using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int DefaultPageWidthTwips = 12240;
        private const int DefaultPageHeightTwips = 15840;
        private const int DefaultPageMarginTwips = 1440;
        private const int DefaultHeaderFooterMarginTwips = 720;
        private const int DefaultColumnSpaceTwips = 720;
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

        private static LegacyDocSectionFormat ReadSupportedSectionProperties(LegacyDocElement sectionProperties) {
            int? pageWidth = null;
            int? pageHeight = null;
            string? orientation = null;
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
            IReadOnlyList<WordSectionColumn>? columnDefinitions = null;
            bool hasColumnSeparator = false;
            int? pageNumberStart = null;
            string? pageNumberFormat = null;
            bool rtlGutter = false;
            string? verticalAlignment = null;
            int? lineNumberCountBy = null;
            int? lineNumberDistance = null;
            int? lineNumberStart = null;
            string? lineNumberRestart = null;
            string? footnotePosition = null;
            string? footnoteRestart = null;
            int? footnoteStart = null;
            string? footnoteNumberFormat = null;
            string? endnotePosition = null;
            string? endnoteRestart = null;
            int? endnoteStart = null;
            string? endnoteNumberFormat = null;
            LegacyDocParagraphBorders? pageBorders = null;
            string? sectionBreakType = null;

            foreach (LegacyDocElement property in sectionProperties.Elements()) {
                switch (property.WName) {
                    case "headerReference":
                    case "footerReference":
                        break;
                    case "titlePg":
                        differentFirstPage = IsOnOffEnabled(property);
                        break;
                    case "pgSz":
                        pageWidth = ReadTwipValue(property.Attr("w"), DefaultPageWidthTwips, "section page width");
                        pageHeight = ReadTwipValue(property.Attr("h"), DefaultPageHeightTwips, "section page height");
                        orientation = property.Attr("orient");
                        if (orientation == null && pageWidth > pageHeight) {
                            orientation = "landscape";
                        }

                        break;
                    case "pgMar":
                        marginTop = ReadTwipValue(property.Attr("top"), DefaultPageMarginTwips, "section top margin");
                        marginRight = ReadTwipValue(property.Attr("right"), DefaultPageMarginTwips, "section right margin");
                        marginBottom = ReadTwipValue(property.Attr("bottom"), DefaultPageMarginTwips, "section bottom margin");
                        marginLeft = ReadTwipValue(property.Attr("left"), DefaultPageMarginTwips, "section left margin");
                        headerDistance = ReadTwipValue(property.Attr("header"), DefaultHeaderFooterMarginTwips, "section header distance");
                        footerDistance = ReadTwipValue(property.Attr("footer"), DefaultHeaderFooterMarginTwips, "section footer distance");
                        gutter = ReadTwipValue(property.Attr("gutter"), 0, "section gutter");
                        break;
                    case "type":
                        if (property.Val != null) {
                            sectionBreakType = property.Val;
                        }

                        break;
                    case "cols":
                        columnDefinitions = ReadSupportedColumns(property, out columnCount, out columnSpacing, out hasColumnSeparator);
                        break;
                    case "pgNumType":
                        pageNumberStart = ReadPageNumberStart(property.Attr("start"));
                        pageNumberFormat = ReadPageNumberFormat(property.Attr("fmt"));
                        break;
                    case "lnNumType":
                        ReadSupportedLineNumbering(property, out lineNumberCountBy, out lineNumberDistance, out lineNumberStart, out lineNumberRestart);
                        break;
                    case "footnotePr":
                        ReadSupportedFootnoteProperties(property, out footnotePosition, out footnoteRestart, out footnoteStart, out footnoteNumberFormat);
                        break;
                    case "endnotePr":
                        ReadSupportedEndnoteProperties(property, out endnotePosition, out endnoteRestart, out endnoteStart, out endnoteNumberFormat);
                        break;
                    case "rtlGutter":
                        rtlGutter = IsOnOffEnabled(property);
                        break;
                    case "vAlign":
                        verticalAlignment = ReadVerticalAlignment(property.Val);
                        break;
                    case "pgBorders":
                        pageBorders = ReadSupportedSectionPageBorders(property);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports simple section page setup only. Unsupported section property: {property.LocalName}.");
                }
            }

            return new LegacyDocSectionFormat(
                ToModel<LegacyDocSectionMark>(sectionBreakType),
                pageWidth,
                pageHeight,
                ToModel<LegacyDocPageOrientation>(orientation),
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
                ToModel<LegacyDocNumberFormat>(pageNumberFormat),
                rtlGutter,
                ToModel<LegacyDocVerticalJustification>(verticalAlignment),
                lineNumberCountBy,
                lineNumberDistance,
                lineNumberStart,
                ToModel<LegacyDocLineNumberRestart>(lineNumberRestart),
                ToModel<LegacyDocFootnotePosition>(footnotePosition),
                ToModel<LegacyDocRestartNumber>(footnoteRestart),
                footnoteStart,
                ToModel<LegacyDocNumberFormat>(footnoteNumberFormat),
                ToModel<LegacyDocEndnotePosition>(endnotePosition),
                ToModel<LegacyDocRestartNumber>(endnoteRestart),
                endnoteStart,
                ToModel<LegacyDocNumberFormat>(endnoteNumberFormat),
                pageBorders,
                columnDefinitions);
        }

        private static LegacyDocParagraphBorders ReadSupportedSectionPageBorders(LegacyDocElement pageBorders) {
            foreach (LegacyDocAttribute attribute in pageBorders.Attributes) {
                if (!string.Equals(attribute.LocalName, "display", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "offsetFrom", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "zOrder", StringComparison.Ordinal)) {
                    throw new NotSupportedException($"Native DOC saving supports simple section page borders only. Unsupported page border attribute: {attribute.LocalName}.");
                }
            }

            foreach (LegacyDocElement child in pageBorders.Elements()) {
                switch (child.WName) {
                    case "top":
                    case "left":
                    case "bottom":
                    case "right":
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple section page borders only. Unsupported page border: {child.LocalName}.");
                }
            }

            return new LegacyDocParagraphBorders(
                ReadSupportedSectionPageBorder(pageBorders.Child("top")),
                ReadSupportedSectionPageBorder(pageBorders.Child("left")),
                ReadSupportedSectionPageBorder(pageBorders.Child("bottom")),
                ReadSupportedSectionPageBorder(pageBorders.Child("right")),
                default,
                ReadSupportedSectionPageBorderOptions(pageBorders));
        }

        private static LegacyDocPageBorderOptions ReadSupportedSectionPageBorderOptions(LegacyDocElement pageBorders) {
            return new LegacyDocPageBorderOptions(
                ReadSupportedSectionPageBorderDisplay(pageBorders.Attr("display")),
                ReadSupportedSectionPageBorderOffsetFrom(pageBorders.Attr("offsetFrom")),
                ReadSupportedSectionPageBorderZOrder(pageBorders.Attr("zOrder")));
        }

        private static LegacyDocPageBorderDisplay ReadSupportedSectionPageBorderDisplay(string? display) {
            if (display == null) {
                return LegacyDocPageBorderDisplay.AllPages;
            }

            if (display == "allPages") {
                return LegacyDocPageBorderDisplay.AllPages;
            }

            if (display == "firstPage") {
                return LegacyDocPageBorderDisplay.FirstPage;
            }

            if (display == "notFirstPage") {
                return LegacyDocPageBorderDisplay.NotFirstPage;
            }

            throw new NotSupportedException($"Native DOC saving does not support section page border display '{FormatSdkEnumValue(display, "PageBorderDisplayValues")}'.");
        }

        private static LegacyDocPageBorderOffsetFrom ReadSupportedSectionPageBorderOffsetFrom(string? offsetFrom) {
            if (offsetFrom == null) {
                return LegacyDocPageBorderOffsetFrom.Text;
            }

            if (offsetFrom == "text") {
                return LegacyDocPageBorderOffsetFrom.Text;
            }

            if (offsetFrom == "page") {
                return LegacyDocPageBorderOffsetFrom.Page;
            }

            throw new NotSupportedException($"Native DOC saving does not support section page border offset '{FormatSdkEnumValue(offsetFrom, "PageBorderOffsetValues")}'.");
        }

        private static LegacyDocPageBorderZOrder ReadSupportedSectionPageBorderZOrder(string? zOrder) {
            if (zOrder == null) {
                return LegacyDocPageBorderZOrder.Front;
            }

            if (zOrder == "front") {
                return LegacyDocPageBorderZOrder.Front;
            }

            if (zOrder == "back") {
                return LegacyDocPageBorderZOrder.Back;
            }

            throw new NotSupportedException($"Native DOC saving does not support section page border z-order '{FormatSdkEnumValue(zOrder, "PageBorderZOrderValues")}'.");
        }

        private static LegacyDocParagraphBorder ReadSupportedSectionPageBorder(LegacyDocElement border) {
            if (border.IsNone) {
                return default;
            }

            foreach (LegacyDocAttribute attribute in border.Attributes) {
                if (!string.Equals(attribute.LocalName, "val", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "color", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "sz", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "space", StringComparison.Ordinal)) {
                    throw new NotSupportedException($"Native DOC saving supports simple section page borders only. Unsupported page border attribute: {attribute.LocalName}.");
                }
            }

            string? value = border.Val;
            if (value == null || value == "none" || value == "nil") {
                return default;
            }

            LegacyDocParagraphBorderStyle style = MapSupportedSectionPageBorderStyle(value);
            string? colorHex = border.Attr("color");
            if (string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                colorHex = null;
            }

            if (!LegacyDocPaletteProjection.CanResolveColor(colorHex)) {
                throw new NotSupportedException("Native DOC saving supports section page borders only with Word 97-2003 palette colors.");
            }

            uint? borderSize = LegacyDocValues.UInt32(border.Attr("sz"));
            int size = borderSize == null ? 4 : checked((int)borderSize.Value);
            uint? borderSpace = LegacyDocValues.UInt32(border.Attr("space"));
            int space = borderSpace == null ? 0 : checked((int)borderSpace.Value);
            if (size <= 0 || size > byte.MaxValue || space < 0 || space > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section page border size and spacing only within Word 97-2003 BRC80 byte ranges.");
            }

            return new LegacyDocParagraphBorder(style, colorHex, size, space);
        }

        private static LegacyDocParagraphBorderStyle MapSupportedSectionPageBorderStyle(string value) {
            if (value == "single") {
                return LegacyDocParagraphBorderStyle.Single;
            }

            if (value == "double") {
                return LegacyDocParagraphBorderStyle.Double;
            }

            if (value == "dotted") {
                return LegacyDocParagraphBorderStyle.Dotted;
            }

            if (value == "dashed" || value == "dashSmallGap") {
                return LegacyDocParagraphBorderStyle.Dashed;
            }

            throw new NotSupportedException($"Native DOC saving does not support section page border style '{FormatSdkEnumValue(value, "BorderValues")}'.");
        }

        private static bool IsOnOffEnabled(LegacyDocElement element) {
            return element.Val is not { } value || LegacyDocValues.ParseOnOff(value);
        }

        private static void ReadSupportedLineNumbering(
            LegacyDocElement lineNumberType,
            out int? countBy,
            out int? distance,
            out int? start,
            out string? restart) {
            countBy = ReadLineNumberCountBy(lineNumberType.Attr("countBy"));
            distance = ReadLineNumberDistance(lineNumberType.Attr("distance"));
            start = ReadLineNumberStart(lineNumberType.Attr("start"));
            restart = ReadLineNumberRestart(lineNumberType.Attr("restart"));
        }

        private static void ReadSupportedFootnoteProperties(
            LegacyDocElement footnoteProperties,
            out string? position,
            out string? restart,
            out int? start,
            out string? numberFormat) {
            position = null;
            restart = null;
            start = null;
            numberFormat = null;

            foreach (LegacyDocElement property in footnoteProperties.Elements()) {
                switch (property.WName) {
                    case "pos":
                        position = ReadFootnotePosition(property.Val);
                        break;
                    case "numRestart":
                        restart = ReadFootnoteRestart(property.Val);
                        break;
                    case "numStart":
                        start = ReadNoteNumberStart(property.Val, "section footnote start");
                        break;
                    case "numFmt":
                        numberFormat = ReadPageNumberFormat(property.Val);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving does not support section footnote property '{property.LocalName}'.");
                }
            }
        }

        private static void ReadSupportedEndnoteProperties(
            LegacyDocElement endnoteProperties,
            out string? position,
            out string? restart,
            out int? start,
            out string? numberFormat) {
            position = null;
            restart = null;
            start = null;
            numberFormat = null;

            foreach (LegacyDocElement property in endnoteProperties.Elements()) {
                switch (property.WName) {
                    case "pos":
                        position = ReadEndnotePosition(property.Val);
                        break;
                    case "numRestart":
                        restart = ReadEndnoteRestart(property.Val);
                        break;
                    case "numStart":
                        start = ReadNoteNumberStart(property.Val, "section endnote start");
                        break;
                    case "numFmt":
                        numberFormat = ReadPageNumberFormat(property.Val);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving does not support section endnote property '{property.LocalName}'.");
                }
            }
        }

        private static int? ReadTwipValue(string? value, int defaultValue, string description) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > ushort.MaxValue) {
                throw new NotSupportedException($"Native DOC saving supports {description} only within the Word 97-2003 unsigned twip range.");
            }

            return actual == defaultValue ? null : actual;
        }

        private static int? ReadPageNumberStart(string? value) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section page number starts only within the Word 97-2003 unsigned range.");
            }

            return actual;
        }

        private static string? ReadPageNumberFormat(string? value) {
            if (value == null) {
                return null;
            }

            return GetPageNumberFormatOperand(value) != null
                ? value
                : throw new NotSupportedException($"Native DOC saving does not support section page number format '{FormatSdkEnumValue(value, "NumberFormatValues")}'.");
        }

        private static string? ReadFootnotePosition(string? value) {
            if (value == null) {
                return null;
            }

            return GetFootnotePositionOperand(value) != null
                ? value
                : throw new NotSupportedException($"Native DOC saving does not support section footnote placement '{FormatSdkEnumValue(value, "FootnotePositionValues")}'.");
        }

        private static string? ReadEndnotePosition(string? value) {
            if (value == null) {
                return null;
            }

            return GetEndnotePositionOperand(value) != null
                ? value
                : throw new NotSupportedException($"Native DOC saving does not support section endnote placement '{FormatSdkEnumValue(value, "EndnotePositionValues")}'.");
        }

        private static string? ReadFootnoteRestart(string? value) {
            if (value == null) {
                return null;
            }

            return GetNoteRestartOperand(value) != null
                ? value
                : throw new NotSupportedException($"Native DOC saving does not support section footnote numbering restart '{FormatSdkEnumValue(value, "RestartNumberValues")}'.");
        }

        private static string? ReadEndnoteRestart(string? value) {
            if (value == null) {
                return null;
            }

            if (value == "eachPage") {
                throw new NotSupportedException("Native DOC saving does not support section endnote numbering restart for each page.");
            }

            return GetNoteRestartOperand(value) != null
                ? value
                : throw new NotSupportedException($"Native DOC saving does not support section endnote numbering restart '{FormatSdkEnumValue(value, "RestartNumberValues")}'.");
        }

        private static int? ReadNoteNumberStart(string? value, string description) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 1
                || actual > 16383) {
                throw new NotSupportedException($"Native DOC saving supports {description} only from 1 through 16383.");
            }

            return actual;
        }

        private static int? ReadLineNumberCountBy(string? value) {
            if (value == null) {
                return 1;
            }

            if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section line number intervals from 0 through the WordprocessingML signed 16-bit range.");
            }

            return actual;
        }

        private static int? ReadLineNumberDistance(string? value) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 0
                || actual > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section line number distance only within the Word 97-2003 unsigned twip range.");
            }

            return actual;
        }

        private static int? ReadLineNumberStart(string? value) {
            if (value == null) {
                return null;
            }

            if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                || actual < 1
                || actual > 32767) {
                throw new NotSupportedException("Native DOC saving supports section line number starts from 1 through 32767.");
            }

            return actual;
        }

        private static string? ReadLineNumberRestart(string? value) {
            if (value == null) {
                return null;
            }

            return GetLineNumberRestartOperand(value) != null
                ? value
                : throw new NotSupportedException($"Native DOC saving does not support section line number restart mode '{FormatSdkEnumValue(value, "LineNumberRestartValues")}'.");
        }

        private static string? ReadVerticalAlignment(string? value) {
            if (value == null || value == "top") {
                return null;
            }

            return GetVerticalAlignmentOperand(value) != null
                ? value
                : throw new NotSupportedException($"Native DOC saving does not support section vertical alignment '{FormatSdkEnumValue(value, "VerticalJustificationValues")}'.");
        }

        private static byte[] CreateSepx(LegacyDocSectionFormat sectionFormat) {
            var grpprl = new List<byte>();

            if (sectionFormat.SectionBreakType != null && sectionFormat.SectionBreakType.Value != LegacyDocSectionMark.NextPage) {
                AddSingleByteSprm(grpprl, SprmSBkc, GetSectionBreakTypeOperand(sectionFormat.SectionBreakType.Value.ToXml()));
            }

            if (sectionFormat.ColumnCount != null) {
                AddUInt16Sprm(grpprl, SprmSCcolumns, sectionFormat.ColumnCount.Value - 1);
            }

            if (sectionFormat.ColumnSpacingTwips != null) {
                AddUInt16Sprm(grpprl, SprmSDxaColumns, sectionFormat.ColumnSpacingTwips.Value);
            }

            AddSectionColumnDefinitions(grpprl, sectionFormat.ColumnDefinitions);

            if (sectionFormat.HasColumnSeparator) {
                AddSingleByteSprm(grpprl, SprmSLBetween, 1);
            }

            if (sectionFormat.PageNumberFormat != null) {
                AddSingleByteSprm(grpprl, SprmSNfcPgn, GetPageNumberFormatOperand(sectionFormat.PageNumberFormat.Value.ToXml())!.Value);
            }

            if (sectionFormat.PageNumberStart != null) {
                AddSingleByteSprm(grpprl, SprmSFPgnRestart, 1);
                AddUInt16Sprm(grpprl, SprmSPgnStart97, sectionFormat.PageNumberStart.Value);
            }

            if (sectionFormat.FootnotePosition != null) {
                AddSingleByteSprm(grpprl, SprmSFpc, GetFootnotePositionOperand(sectionFormat.FootnotePosition.Value.ToXml())!.Value);
            }

            if (sectionFormat.FootnoteRestart != null) {
                AddSingleByteSprm(grpprl, SprmSRncFtn, GetNoteRestartOperand(sectionFormat.FootnoteRestart.Value.ToXml())!.Value);
            }

            if (sectionFormat.EndnoteRestart != null) {
                AddSingleByteSprm(grpprl, SprmSRncEdn, GetNoteRestartOperand(sectionFormat.EndnoteRestart.Value.ToXml())!.Value);
            }

            if (sectionFormat.FootnoteStart != null) {
                AddUInt16Sprm(grpprl, SprmSNFtn, sectionFormat.FootnoteStart.Value);
            }

            if (sectionFormat.FootnoteNumberFormat != null) {
                AddUInt16Sprm(grpprl, SprmSNfcFtnRef, GetPageNumberFormatOperand(sectionFormat.FootnoteNumberFormat.Value.ToXml())!.Value);
            }

            if (sectionFormat.EndnoteStart != null) {
                AddUInt16Sprm(grpprl, SprmSNEdn, sectionFormat.EndnoteStart.Value);
            }

            if (sectionFormat.EndnoteNumberFormat != null) {
                AddUInt16Sprm(grpprl, SprmSNfcEdnRef, GetPageNumberFormatOperand(sectionFormat.EndnoteNumberFormat.Value.ToXml())!.Value);
            }

            if (sectionFormat.LineNumberRestart != null) {
                AddSingleByteSprm(grpprl, SprmSLnc, GetLineNumberRestartOperand(sectionFormat.LineNumberRestart.Value.ToXml())!.Value);
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

            if (sectionFormat.Orientation == LegacyDocPageOrientation.Landscape) {
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
                AddSingleByteSprm(grpprl, SprmSVjc, GetVerticalAlignmentOperand(sectionFormat.VerticalAlignment.Value.ToXml())!.Value);
            }

            if (sectionFormat.PageBorders != null) {
                AddSectionPageBorderSprms(grpprl, sectionFormat.PageBorders.Value);
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

        private static void AddSectionPageBorderSprms(List<byte> grpprl, LegacyDocParagraphBorders borders) {
            AddSectionPageBorderSprmIfPresent(grpprl, SprmSBrcTop80, borders.Top);
            AddSectionPageBorderSprmIfPresent(grpprl, SprmSBrcLeft80, borders.Left);
            AddSectionPageBorderSprmIfPresent(grpprl, SprmSBrcBottom80, borders.Bottom);
            AddSectionPageBorderSprmIfPresent(grpprl, SprmSBrcRight80, borders.Right);

            if (borders.PageOptions.HasNonDefault) {
                AddSectionPageBorderOptionsSprm(grpprl, borders.PageOptions);
            }
        }

        private static void AddSectionPageBorderOptionsSprm(List<byte> grpprl, LegacyDocPageBorderOptions options) {
            byte operand = (byte)(GetSectionPageBorderDisplayOperand(options.Display)
                | (GetSectionPageBorderZOrderOperand(options.ZOrder) << 3)
                | (GetSectionPageBorderOffsetOperand(options.OffsetFrom) << 5));

            grpprl.Add((byte)(SprmSPgbProp & 0xFF));
            grpprl.Add((byte)(SprmSPgbProp >> 8));
            grpprl.Add(operand);
            grpprl.Add(0);
        }

        private static void AddSectionPageBorderSprmIfPresent(List<byte> grpprl, ushort sprm, LegacyDocParagraphBorder border) {
            if (!border.HasAny) {
                return;
            }

            AddSectionPageBorderSprm(grpprl, sprm, border);
        }

        private static void AddSectionPageBorderSprm(List<byte> grpprl, ushort sprm, LegacyDocParagraphBorder border) {
            if (border.SizeEighthPoints <= 0 || border.SizeEighthPoints > byte.MaxValue || border.SpacePoints < 0 || border.SpacePoints > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports section page border size and spacing only within Word 97-2003 BRC80 byte ranges.");
            }

            if (!TryMapSectionPageBorderStyle(border.Style, out byte borderType)) {
                throw new NotSupportedException($"Native DOC saving does not support section page border style '{border.Style}'.");
            }

            if (!LegacyDocPaletteProjection.TryResolveColor(border.ColorHex, "page border", out byte colorIndex)) {
                throw new NotSupportedException("Native DOC saving supports section page borders only with Word 97-2003 palette colors.");
            }

            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add(checked((byte)border.SizeEighthPoints));
            grpprl.Add(borderType);
            grpprl.Add(colorIndex);
            grpprl.Add(checked((byte)border.SpacePoints));
        }

        private static bool TryMapSectionPageBorderStyle(LegacyDocParagraphBorderStyle style, out byte borderType) {
            switch (style) {
                case LegacyDocParagraphBorderStyle.Single:
                    borderType = 0x01;
                    return true;
                case LegacyDocParagraphBorderStyle.Double:
                    borderType = 0x03;
                    return true;
                case LegacyDocParagraphBorderStyle.Dotted:
                    borderType = 0x06;
                    return true;
                case LegacyDocParagraphBorderStyle.Dashed:
                    borderType = 0x07;
                    return true;
                default:
                    borderType = 0;
                    return false;
            }
        }

        private static byte GetSectionPageBorderDisplayOperand(LegacyDocPageBorderDisplay display) {
            switch (display) {
                case LegacyDocPageBorderDisplay.AllPages:
                    return 0x00;
                case LegacyDocPageBorderDisplay.FirstPage:
                    return 0x01;
                case LegacyDocPageBorderDisplay.NotFirstPage:
                    return 0x02;
                default:
                    throw new NotSupportedException($"Native DOC saving does not support section page border display '{display}'.");
            }
        }

        private static byte GetSectionPageBorderOffsetOperand(LegacyDocPageBorderOffsetFrom offsetFrom) {
            switch (offsetFrom) {
                case LegacyDocPageBorderOffsetFrom.Text:
                    return 0x00;
                case LegacyDocPageBorderOffsetFrom.Page:
                    return 0x01;
                default:
                    throw new NotSupportedException($"Native DOC saving does not support section page border offset '{offsetFrom}'.");
            }
        }

        private static byte GetSectionPageBorderZOrderOperand(LegacyDocPageBorderZOrder zOrder) {
            switch (zOrder) {
                case LegacyDocPageBorderZOrder.Front:
                    return 0x00;
                case LegacyDocPageBorderZOrder.Back:
                    return 0x01;
                default:
                    throw new NotSupportedException($"Native DOC saving does not support section page border z-order '{zOrder}'.");
            }
        }

        private static byte GetSectionBreakTypeOperand(string sectionBreakType) {
            if (sectionBreakType == "continuous") {
                return 0;
            }

            if (sectionBreakType == "nextColumn") {
                return 1;
            }

            if (sectionBreakType == "nextPage") {
                return 2;
            }

            if (sectionBreakType == "evenPage") {
                return 3;
            }

            if (sectionBreakType == "oddPage") {
                return 4;
            }

            throw new NotSupportedException($"Native DOC saving does not support section break type '{FormatSdkEnumValue(sectionBreakType, "SectionMarkValues")}'.");
        }

        private static byte? GetPageNumberFormatOperand(string format) {
            return ToNfc(format);
        }

        private static byte? GetLineNumberRestartOperand(string restart) {
            if (restart == "newPage") {
                return 0;
            }

            if (restart == "newSection") {
                return 1;
            }

            if (restart == "continuous") {
                return 2;
            }

            return null;
        }

        private static byte? GetFootnotePositionOperand(string position) {
            if (position == "pageBottom") {
                return 1;
            }

            if (position == "beneathText") {
                return 2;
            }

            return null;
        }

        private static byte? GetEndnotePositionOperand(string position) {
            if (position == "sectEnd") {
                return 0;
            }

            if (position == "docEnd") {
                return 3;
            }

            return null;
        }

        private static byte? GetNoteRestartOperand(string restart) {
            if (restart == "continuous") {
                return 0;
            }

            if (restart == "eachSect") {
                return 1;
            }

            if (restart == "eachPage") {
                return 2;
            }

            return null;
        }

        private static byte? GetVerticalAlignmentOperand(string alignment) {
            if (alignment == "top") {
                return 0;
            }

            if (alignment == "center") {
                return 1;
            }

            if (alignment == "both") {
                return 2;
            }

            if (alignment == "bottom") {
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
