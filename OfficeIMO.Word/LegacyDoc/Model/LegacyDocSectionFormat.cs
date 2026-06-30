using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocSectionFormat {
        internal LegacyDocSectionFormat(
            SectionMarkValues? sectionBreakType,
            int? pageWidthTwips,
            int? pageHeightTwips,
            PageOrientationValues? orientation,
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
            NumberFormatValues? pageNumberFormat = null,
            bool rtlGutter = false,
            VerticalJustificationValues? verticalAlignment = null,
            int? lineNumberCountBy = null,
            int? lineNumberDistanceTwips = null,
            int? lineNumberStart = null,
            LineNumberRestartValues? lineNumberRestart = null) {
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
            HasColumnSeparator = hasColumnSeparator;
            PageNumberStart = pageNumberStart;
            PageNumberFormat = pageNumberFormat;
            RtlGutter = rtlGutter;
            VerticalAlignment = verticalAlignment;
            LineNumberCountBy = lineNumberCountBy;
            LineNumberDistanceTwips = lineNumberDistanceTwips;
            LineNumberStart = lineNumberStart;
            LineNumberRestart = lineNumberRestart;
        }

        internal SectionMarkValues? SectionBreakType { get; }

        internal int? PageWidthTwips { get; }

        internal int? PageHeightTwips { get; }

        internal PageOrientationValues? Orientation { get; }

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

        internal bool HasColumnSeparator { get; }

        internal int? PageNumberStart { get; }

        internal NumberFormatValues? PageNumberFormat { get; }

        internal bool RtlGutter { get; }

        internal VerticalJustificationValues? VerticalAlignment { get; }

        internal int? LineNumberCountBy { get; }

        internal int? LineNumberDistanceTwips { get; }

        internal int? LineNumberStart { get; }

        internal LineNumberRestartValues? LineNumberRestart { get; }

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
            || HasColumnSeparator
            || PageNumberStart != null
            || PageNumberFormat != null
            || RtlGutter
            || VerticalAlignment != null
            || LineNumberCountBy != null
            || LineNumberDistanceTwips != null
            || LineNumberStart != null
            || LineNumberRestart != null;

        private static bool IsNonDefaultSectionBreakType(SectionMarkValues? sectionBreakType) {
            return sectionBreakType != null && sectionBreakType.Value != SectionMarkValues.NextPage;
        }

        internal LegacyDocSectionFormat WithSectionBreakType(SectionMarkValues? sectionBreakType) {
            return new LegacyDocSectionFormat(
                sectionBreakType,
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
                LineNumberRestart);
        }

        internal static LegacyDocSectionFormat Default { get; } = new LegacyDocSectionFormat(null, null, null, null, null, null, null, null, null, null, null);
    }
}
