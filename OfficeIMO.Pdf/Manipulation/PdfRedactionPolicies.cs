namespace OfficeIMO.Pdf;

/// <summary>Controls which intersecting page content is removed for one redaction area.</summary>
public enum PdfRedactionContentScope {
    /// <summary>Remove intersecting text and matched annotations while preserving images and vector artwork beneath the mark.</summary>
    TextOnly,
    /// <summary>Remove intersecting text, matched annotations, images, and vector artwork.</summary>
    TextAndUnderlay
}

/// <summary>Controls how the visible redaction mark conceals the size of the removed content.</summary>
public enum PdfRedactionAppearanceMode {
    /// <summary>Paint the exact reviewed geometry.</summary>
    Exact,
    /// <summary>Merge nearby rectangular marks on the same visual line.</summary>
    MergeNearby,
    /// <summary>Expand rectangular marks to a fixed width quantum.</summary>
    QuantizedWidth,
    /// <summary>Paint the full effective page width at the reviewed vertical position.</summary>
    FullLine
}
