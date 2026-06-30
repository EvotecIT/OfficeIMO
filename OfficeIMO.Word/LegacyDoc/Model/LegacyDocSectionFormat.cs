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
            bool hasColumnSeparator = false) {
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
            || HasColumnSeparator;

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
                HasColumnSeparator);
        }

        internal static LegacyDocSectionFormat Default { get; } = new LegacyDocSectionFormat(null, null, null, null, null, null, null, null, null, null, null);
    }
}
