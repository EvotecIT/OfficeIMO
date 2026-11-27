
namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocSectionFormat {
        internal LegacyDocSectionFormat(
            LegacyDocSectionMark? sectionBreakType,
            int? pageWidthTwips,
            int? pageHeightTwips,
            LegacyDocPageOrientation? orientation,
            int? marginTopTwips,
            int? marginRightTwips,
            int? marginBottomTwips,
            int? marginLeftTwips,
            int? headerDistanceTwips,
            int? footerDistanceTwips,
            int? gutterTwips,
            bool differentFirstPage = false,
            int? columnCount = null,
            int? columnSpacingTwips = null,
            bool hasColumnSeparator = false,
            int? pageNumberStart = null,
            LegacyDocNumberFormat? pageNumberFormat = null,
            bool rtlGutter = false,
            LegacyDocVerticalJustification? verticalAlignment = null,
            int? lineNumberCountBy = null,
            int? lineNumberDistanceTwips = null,
            int? lineNumberStart = null,
            LegacyDocLineNumberRestart? lineNumberRestart = null,
            LegacyDocFootnotePosition? footnotePosition = null,
            LegacyDocRestartNumber? footnoteRestart = null,
            int? footnoteStart = null,
            LegacyDocNumberFormat? footnoteNumberFormat = null,
            LegacyDocEndnotePosition? endnotePosition = null,
            LegacyDocRestartNumber? endnoteRestart = null,
            int? endnoteStart = null,
            LegacyDocNumberFormat? endnoteNumberFormat = null,
            LegacyDocParagraphBorders? pageBorders = null,
            IReadOnlyList<WordSectionColumn>? columnDefinitions = null) {
            SectionBreakType = sectionBreakType;
            PageWidthTwips = pageWidthTwips;
            PageHeightTwips = pageHeightTwips;
            Orientation = orientation;
            MarginTopTwips = marginTopTwips;
            MarginRightTwips = marginRightTwips;
            MarginBottomTwips = marginBottomTwips;
            MarginLeftTwips = marginLeftTwips;
            HeaderDistanceTwips = headerDistanceTwips;
            FooterDistanceTwips = footerDistanceTwips;
            GutterTwips = gutterTwips;
            DifferentFirstPage = differentFirstPage;
            ColumnCount = columnCount;
            ColumnSpacingTwips = columnSpacingTwips;
            ColumnDefinitions = columnDefinitions == null ? null : Array.AsReadOnly(columnDefinitions.ToArray());
            HasColumnSeparator = hasColumnSeparator;
            PageNumberStart = pageNumberStart;
            PageNumberFormat = pageNumberFormat;
            RtlGutter = rtlGutter;
            VerticalAlignment = verticalAlignment;
            LineNumberCountBy = lineNumberCountBy;
            LineNumberDistanceTwips = lineNumberDistanceTwips;
            LineNumberStart = lineNumberStart;
            LineNumberRestart = lineNumberRestart;
            FootnotePosition = footnotePosition;
            FootnoteRestart = footnoteRestart;
            FootnoteStart = footnoteStart;
            FootnoteNumberFormat = footnoteNumberFormat;
            EndnotePosition = endnotePosition;
            EndnoteRestart = endnoteRestart;
            EndnoteStart = endnoteStart;
            EndnoteNumberFormat = endnoteNumberFormat;
            PageBorders = pageBorders.HasValue && pageBorders.Value.HasAny ? pageBorders : null;
        }

        internal LegacyDocSectionMark? SectionBreakType { get; }

        internal int? PageWidthTwips { get; }

        internal int? PageHeightTwips { get; }

        internal LegacyDocPageOrientation? Orientation { get; }

        internal int? MarginTopTwips { get; }

        internal int? MarginRightTwips { get; }

        internal int? MarginBottomTwips { get; }

        internal int? MarginLeftTwips { get; }

        internal int? HeaderDistanceTwips { get; }

        internal int? FooterDistanceTwips { get; }

        internal int? GutterTwips { get; }

        internal bool DifferentFirstPage { get; }

        internal int? ColumnCount { get; }

        internal int? ColumnSpacingTwips { get; }

        internal IReadOnlyList<WordSectionColumn>? ColumnDefinitions { get; }

        internal bool HasColumnSeparator { get; }

        internal int? PageNumberStart { get; }

        internal LegacyDocNumberFormat? PageNumberFormat { get; }

        internal bool RtlGutter { get; }

        internal LegacyDocVerticalJustification? VerticalAlignment { get; }

        internal int? LineNumberCountBy { get; }

        internal int? LineNumberDistanceTwips { get; }

        internal int? LineNumberStart { get; }

        internal LegacyDocLineNumberRestart? LineNumberRestart { get; }

        internal LegacyDocFootnotePosition? FootnotePosition { get; }

        internal LegacyDocRestartNumber? FootnoteRestart { get; }

        internal int? FootnoteStart { get; }

        internal LegacyDocNumberFormat? FootnoteNumberFormat { get; }

        internal LegacyDocEndnotePosition? EndnotePosition { get; }

        internal LegacyDocRestartNumber? EndnoteRestart { get; }

        internal int? EndnoteStart { get; }

        internal LegacyDocNumberFormat? EndnoteNumberFormat { get; }

        internal LegacyDocParagraphBorders? PageBorders { get; }

        internal bool HasFormatting => IsNonDefaultSectionBreakType(SectionBreakType)
            || PageWidthTwips != null
            || PageHeightTwips != null
            || Orientation != null
            || MarginTopTwips != null
            || MarginRightTwips != null
            || MarginBottomTwips != null
            || MarginLeftTwips != null
            || HeaderDistanceTwips != null
            || FooterDistanceTwips != null
            || GutterTwips != null
            || DifferentFirstPage
            || ColumnCount != null
            || ColumnSpacingTwips != null
            || ColumnDefinitions?.Count > 0
            || HasColumnSeparator
            || PageNumberStart != null
            || PageNumberFormat != null
            || RtlGutter
            || VerticalAlignment != null
            || LineNumberCountBy != null
            || LineNumberDistanceTwips != null
            || LineNumberStart != null
            || LineNumberRestart != null
            || FootnotePosition != null
            || FootnoteRestart != null
            || FootnoteStart != null
            || FootnoteNumberFormat != null
            || EndnotePosition != null
            || EndnoteRestart != null
            || EndnoteStart != null
            || EndnoteNumberFormat != null
            || PageBorders != null;

        private static bool IsNonDefaultSectionBreakType(LegacyDocSectionMark? sectionBreakType) {
            return sectionBreakType != null && sectionBreakType.Value != LegacyDocSectionMark.NextPage;
        }

        internal LegacyDocSectionFormat WithEndnotePosition(LegacyDocEndnotePosition? endnotePosition) {
            return WithNoteSettings(endnotePosition: endnotePosition);
        }

        internal LegacyDocSectionFormat WithNoteSettings(
            LegacyDocFootnotePosition? footnotePosition = null,
            LegacyDocRestartNumber? footnoteRestart = null,
            int? footnoteStart = null,
            LegacyDocNumberFormat? footnoteNumberFormat = null,
            LegacyDocEndnotePosition? endnotePosition = null,
            LegacyDocRestartNumber? endnoteRestart = null,
            int? endnoteStart = null,
            LegacyDocNumberFormat? endnoteNumberFormat = null) {
            return new LegacyDocSectionFormat(
                SectionBreakType,
                PageWidthTwips,
                PageHeightTwips,
                Orientation,
                MarginTopTwips,
                MarginRightTwips,
                MarginBottomTwips,
                MarginLeftTwips,
                HeaderDistanceTwips,
                FooterDistanceTwips,
                GutterTwips,
                DifferentFirstPage,
                ColumnCount,
                ColumnSpacingTwips,
                HasColumnSeparator,
                PageNumberStart,
                PageNumberFormat,
                RtlGutter,
                VerticalAlignment,
                LineNumberCountBy,
                LineNumberDistanceTwips,
                LineNumberStart,
                LineNumberRestart,
                footnotePosition ?? FootnotePosition,
                footnoteRestart ?? FootnoteRestart,
                footnoteStart ?? FootnoteStart,
                footnoteNumberFormat ?? FootnoteNumberFormat,
                endnotePosition ?? EndnotePosition,
                endnoteRestart ?? EndnoteRestart,
                endnoteStart ?? EndnoteStart,
                endnoteNumberFormat ?? EndnoteNumberFormat,
                PageBorders,
                ColumnDefinitions);
        }

        internal static LegacyDocSectionFormat Default { get; } = new LegacyDocSectionFormat(null, null, null, null, null, null, null, null, null, null, null);
    }
}
