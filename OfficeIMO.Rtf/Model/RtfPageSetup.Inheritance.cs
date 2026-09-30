namespace OfficeIMO.Rtf;

public sealed partial class RtfPageSetup {
    internal RtfPageSetup WithFallback(RtfPageSetup fallback) {
        var result = new RtfCloneContext().Clone(this)!;
        result.PaperWidthTwips ??= fallback.PaperWidthTwips;
        result.PaperHeightTwips ??= fallback.PaperHeightTwips;
        result.PrinterPaperSize ??= fallback.PrinterPaperSize;
        result.FirstPagePaperSource ??= fallback.FirstPagePaperSource;
        result.OtherPagesPaperSource ??= fallback.OtherPagesPaperSource;
        result.MarginLeftTwips ??= fallback.MarginLeftTwips;
        result.MarginRightTwips ??= fallback.MarginRightTwips;
        result.MarginTopTwips ??= fallback.MarginTopTwips;
        result.MarginBottomTwips ??= fallback.MarginBottomTwips;
        result.GutterWidthTwips ??= fallback.GutterWidthTwips;
        result.HeaderDistanceTwips ??= fallback.HeaderDistanceTwips;
        result.FooterDistanceTwips ??= fallback.FooterDistanceTwips;
        result.PageNumberStart ??= fallback.PageNumberStart;
        result.PageNumberRestart ??= fallback.PageNumberRestart;
        result.PageNumberPositionXTwips ??= fallback.PageNumberPositionXTwips;
        result.PageNumberPositionYTwips ??= fallback.PageNumberPositionYTwips;
        result.PageNumberFormat ??= fallback.PageNumberFormat;
        result.DirectLandscape ??= fallback.DirectLandscape;
        result.DirectDifferentFirstPageHeaderFooter ??= fallback.DirectDifferentFirstPageHeaderFooter;
        result.DirectRtlGutter ??= fallback.DirectRtlGutter;
        if (!result.PageBorders.HasAnyValue) result.PageBorders = new RtfCloneContext().Clone(fallback.PageBorders)!;
        return result;
    }
}

public sealed partial class RtfDocument {
    /// <summary>Returns an independent page setup with document defaults applied beneath the section's authored values.</summary>
    public RtfPageSetup GetEffectivePageSetup(RtfSection section) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        if (!_sections.Contains(section)) throw new ArgumentException("Section must belong to this document.", nameof(section));
        return GetPageSetup(section);
    }

    internal RtfPageSetup GetPageSetup(RtfSection section) => section.PageSetup.WithFallback(PageSetup);
}
