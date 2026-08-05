namespace OfficeIMO.Drawing;

/// <summary>
/// Base type for ordered elements inside an <see cref="OfficeDrawing"/> canvas.
/// </summary>
public abstract class OfficeDrawingElement {
    internal abstract OfficeDrawingElement CloneElement();
}
