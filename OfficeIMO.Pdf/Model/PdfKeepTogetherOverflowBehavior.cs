namespace OfficeIMO.Pdf;

/// <summary>Controls how imported keep or widow constraints yield when their required lines cannot fit any full frame.</summary>
internal enum PdfKeepTogetherOverflowBehavior {
    Reject,
    SplitAcrossFrames,
    SplitAcrossPages
}
