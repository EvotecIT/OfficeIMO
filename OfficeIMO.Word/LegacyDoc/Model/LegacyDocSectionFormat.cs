using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocSectionFormat {
        internal LegacyDocSectionFormat(
            int? pageWidthTwips,
            int? pageHeightTwips,
            PageOrientationValues? orientation,
            int? marginTopTwips,
            int? marginRightTwips,
            int? marginBottomTwips,
            int? marginLeftTwips,
            int? headerDistanceTwips,
            int? footerDistanceTwips,
            int? gutterTwips) {
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
        }

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

        internal bool HasFormatting => PageWidthTwips != null
            || PageHeightTwips != null
            || Orientation != null
            || MarginTopTwips != null
            || MarginRightTwips != null
            || MarginBottomTwips != null
            || MarginLeftTwips != null
            || HeaderDistanceTwips != null
            || FooterDistanceTwips != null
            || GutterTwips != null;

        internal static LegacyDocSectionFormat Default { get; } = new LegacyDocSectionFormat(null, null, null, null, null, null, null, null, null, null);
    }
}
