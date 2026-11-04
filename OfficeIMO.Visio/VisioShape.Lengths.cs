namespace OfficeIMO.Visio;

public partial class VisioShape {
    /// <summary>Creates geometry from explicit physical lengths.</summary>
    public VisioShape(string id, OfficeLength pinX, OfficeLength pinY, OfficeLength width, OfficeLength height, string text) : this(id, pinX.ToInches(), pinY.ToInches(), width.ToInches(), height.ToInches(), text) { }

    /// <summary>Gets or sets the physical Width.</summary>
    public OfficeLength Width { get => OfficeLength.Inches(WidthInches); set => WidthInches = value.ToInches(); }

    /// <summary>Gets or sets the physical Height.</summary>
    public OfficeLength Height { get => OfficeLength.Inches(HeightInches); set => HeightInches = value.ToInches(); }

    /// <summary>Gets or sets the physical PinX.</summary>
    public OfficeLength PinX { get => OfficeLength.Inches(PinXInches); set => PinXInches = value.ToInches(); }

    /// <summary>Gets or sets the physical PinY.</summary>
    public OfficeLength PinY { get => OfficeLength.Inches(PinYInches); set => PinYInches = value.ToInches(); }

    /// <summary>Gets or sets the physical LocPinX.</summary>
    public OfficeLength LocPinX { get => OfficeLength.Inches(LocPinXInches); set => LocPinXInches = value.ToInches(); }

    /// <summary>Gets or sets the physical LocPinY.</summary>
    public OfficeLength LocPinY { get => OfficeLength.Inches(LocPinYInches); set => LocPinYInches = value.ToInches(); }
    /// <summary>Gets the physical bounds in page coordinates, including rotation and parent transforms.</summary>
    public (OfficeLength Left, OfficeLength Bottom, OfficeLength Right, OfficeLength Top) GetBounds() {
        var bounds = GetBoundsInches();
        return (OfficeLength.Inches(bounds.Left), OfficeLength.Inches(bounds.Bottom), OfficeLength.Inches(bounds.Right), OfficeLength.Inches(bounds.Top));
    }
}
