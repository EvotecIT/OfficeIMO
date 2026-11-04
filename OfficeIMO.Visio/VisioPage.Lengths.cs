namespace OfficeIMO.Visio;

public partial class VisioPage {
    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddContainer(string id, OfficeLength pinX, OfficeLength pinY, OfficeLength width, OfficeLength height, string? text = null, VisioContainerOptions? options = null) => AddContainerInches(id, pinX.ToInches(), pinY.ToInches(), width.ToInches(), height.ToInches(), text, options);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioPage SetMargins(OfficeLength margin) => SetMargins(margin, margin, margin, margin);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioPage SetMargins(OfficeLength horizontal, OfficeLength vertical) => SetMargins(horizontal, horizontal, vertical, vertical);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioPage SetMargins(OfficeLength left, OfficeLength right, OfficeLength top, OfficeLength bottom) {
        SetMarginsInches(left.ToInches(), right.ToInches(), top.ToInches(), bottom.ToInches(), VisioMeasurementUnit.Inches);
        _marginUnit = DefaultUnit;
        return this;
    }

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddShape(string id, VisioMaster master, OfficeLength x, OfficeLength y, OfficeLength w, OfficeLength h, string? text = null) => AddShapeInches(id, master, x.ToInches(), y.ToInches(), w.ToInches(), h.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddShape(string id, string masterNameU, OfficeLength x, OfficeLength y, OfficeLength w, OfficeLength h, string? text = null) => AddShapeInches(id, masterNameU, x.ToInches(), y.ToInches(), w.ToInches(), h.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddRectangle(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddRectangleInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddTextBox(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddTextBoxInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddTextBox(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddTextBoxInches(id, x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddProcess(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddProcessInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddSquare(OfficeLength x, OfficeLength y, OfficeLength size, string? text = null) => AddSquareInches(x.ToInches(), y.ToInches(), size.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddCircle(OfficeLength x, OfficeLength y, OfficeLength diameter, string? text = null) => AddCircleInches(x.ToInches(), y.ToInches(), diameter.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddEllipse(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddEllipseInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddDiamond(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddDiamondInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddDecision(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddDecisionInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddData(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddDataInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddPreparation(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddPreparationInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddParallelogram(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddParallelogramInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddHexagon(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddHexagonInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddTrapezoid(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddTrapezoidInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddPentagon(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddPentagonInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddManualOperation(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddManualOperationInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddOffPageReference(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddOffPageReferenceInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioShape AddTriangle(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddTriangleInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches);

    /// <summary>Applies physical lengths to the native Visio geometry.</summary>
    public VisioPage Size(OfficeLength w, OfficeLength h) => SizeInches(w.ToInches(), h.ToInches(), VisioMeasurementUnit.Inches);

    /// <summary>Creates geometry from explicit physical lengths.</summary>
    public VisioPage(string name, OfficeLength width, OfficeLength height) : this(name, width.ToInches(), height.ToInches()) { }

    /// <summary>Gets or sets the physical Width.</summary>
    public OfficeLength Width { get => OfficeLength.Inches(WidthInches); set => WidthInches = value.ToInches(); }

    /// <summary>Gets or sets the physical Height.</summary>
    public OfficeLength Height { get => OfficeLength.Inches(HeightInches); set => HeightInches = value.ToInches(); }

    /// <summary>Gets the physical LeftMargin.</summary>
    public OfficeLength LeftMargin { get => OfficeLength.Inches(LeftMarginInches); }

    /// <summary>Gets the physical RightMargin.</summary>
    public OfficeLength RightMargin { get => OfficeLength.Inches(RightMarginInches); }

    /// <summary>Gets the physical TopMargin.</summary>
    public OfficeLength TopMargin { get => OfficeLength.Inches(TopMarginInches); }

    /// <summary>Gets the physical BottomMargin.</summary>
    public OfficeLength BottomMargin { get => OfficeLength.Inches(BottomMarginInches); }

    /// <summary>Adds a rectangle with an explicit identifier and physical bounds.</summary>
    public VisioShape AddRectangle(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddRectangleInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a process with an explicit identifier and physical bounds.</summary>
    public VisioShape AddProcess(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddProcessInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a square with an explicit identifier and physical bounds.</summary>
    public VisioShape AddSquare(string id, OfficeLength x, OfficeLength y, OfficeLength size, string? text = null) => AddSquareInches(x.ToInches(), y.ToInches(), size.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a circle with an explicit identifier and physical bounds.</summary>
    public VisioShape AddCircle(string id, OfficeLength x, OfficeLength y, OfficeLength diameter, string? text = null) => AddCircleInches(x.ToInches(), y.ToInches(), diameter.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a ellipse with an explicit identifier and physical bounds.</summary>
    public VisioShape AddEllipse(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddEllipseInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a diamond with an explicit identifier and physical bounds.</summary>
    public VisioShape AddDiamond(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddDiamondInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a decision with an explicit identifier and physical bounds.</summary>
    public VisioShape AddDecision(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddDecisionInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a data with an explicit identifier and physical bounds.</summary>
    public VisioShape AddData(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddDataInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a preparation with an explicit identifier and physical bounds.</summary>
    public VisioShape AddPreparation(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddPreparationInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a parallelogram with an explicit identifier and physical bounds.</summary>
    public VisioShape AddParallelogram(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddParallelogramInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a hexagon with an explicit identifier and physical bounds.</summary>
    public VisioShape AddHexagon(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddHexagonInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a trapezoid with an explicit identifier and physical bounds.</summary>
    public VisioShape AddTrapezoid(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddTrapezoidInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a pentagon with an explicit identifier and physical bounds.</summary>
    public VisioShape AddPentagon(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddPentagonInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a manualoperation with an explicit identifier and physical bounds.</summary>
    public VisioShape AddManualOperation(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddManualOperationInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a offpagereference with an explicit identifier and physical bounds.</summary>
    public VisioShape AddOffPageReference(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddOffPageReferenceInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);

    /// <summary>Adds a triangle with an explicit identifier and physical bounds.</summary>
    public VisioShape AddTriangle(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) => AddTriangleInches(x.ToInches(), y.ToInches(), width.ToInches(), height.ToInches(), text, VisioMeasurementUnit.Inches, id);
}
