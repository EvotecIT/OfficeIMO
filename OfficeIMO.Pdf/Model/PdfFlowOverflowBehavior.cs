namespace OfficeIMO.Pdf;

/// <summary>Controls how a bounded flow group behaves when it cannot fit in the current page remainder.</summary>
public enum PdfFlowOverflowBehavior {
    /// <summary>Allows nested content to continue through the normal page flow.</summary>
    Continue,
    /// <summary>Moves the group to the next page when its measured content fits on one full page.</summary>
    MoveToNextPage,
    /// <summary>Skips the group when it cannot fit in the current page remainder.</summary>
    Skip,
    /// <summary>Stops processing the enclosing document flow when the group cannot fit.</summary>
    StopDocument
}
