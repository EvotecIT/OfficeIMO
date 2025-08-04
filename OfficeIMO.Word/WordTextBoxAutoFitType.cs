namespace OfficeIMO.Word;

/// <summary>
/// Defines the AutoFit options for text within a textbox.
/// </summary>
public enum WordTextBoxAutoFitType {
    /// <summary>
    /// Do not fit text automatically.
    /// </summary>
    NoAutoFit,

    /// <summary>
    /// Shrink text on overflow.
    /// </summary>
    ShrinkTextOnOverflow,

    /// <summary>
    /// Resize the shape to fit the text.
    /// </summary>
    ResizeShapeToFitText
}