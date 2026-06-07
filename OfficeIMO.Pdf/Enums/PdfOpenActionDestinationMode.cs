namespace OfficeIMO.Pdf;

/// <summary>
/// Viewer destination mode used by generated catalog open actions.
/// </summary>
public enum PdfOpenActionDestinationMode {
    /// <summary>Open the target page at a top coordinate using the viewer's current zoom.</summary>
    Xyz = 0,
    /// <summary>Fit the entire target page in the viewer window.</summary>
    Fit = 1,
    /// <summary>Fit the target page horizontally at a top coordinate.</summary>
    FitHorizontal = 2,
    /// <summary>Fit the target page vertically at a left coordinate.</summary>
    FitVertical = 3,
    /// <summary>Fit the specified rectangle in the viewer window.</summary>
    FitRectangle = 4,
    /// <summary>Fit the page bounding box in the viewer window.</summary>
    FitBoundingBox = 5,
    /// <summary>Fit the page bounding box horizontally at a top coordinate.</summary>
    FitBoundingBoxHorizontal = 6,
    /// <summary>Fit the page bounding box vertically at a left coordinate.</summary>
    FitBoundingBoxVertical = 7
}
