namespace OfficeIMO.Visio;

public partial class VisioPage {
    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddRectangle(double x, double y, double width, double height) =>
        AddRectangle(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddTextBox(string id, double x, double y, double width, double height) =>
        AddTextBox(id, x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddTextBox(double x, double y, double width, double height) =>
        AddTextBox(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddProcess(double x, double y, double width, double height) =>
        AddProcess(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddSquare(double x, double y, double size) =>
        AddSquare(x, y, size, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddCircle(double x, double y, double diameter) =>
        AddCircle(x, y, diameter, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddEllipse(double x, double y, double width, double height) =>
        AddEllipse(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddDiamond(double x, double y, double width, double height) =>
        AddDiamond(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddDecision(double x, double y, double width, double height) =>
        AddDecision(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddData(double x, double y, double width, double height) =>
        AddData(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddPreparation(double x, double y, double width, double height) =>
        AddPreparation(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddParallelogram(double x, double y, double width, double height) =>
        AddParallelogram(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddHexagon(double x, double y, double width, double height) =>
        AddHexagon(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddTrapezoid(double x, double y, double width, double height) =>
        AddTrapezoid(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddPentagon(double x, double y, double width, double height) =>
        AddPentagon(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddManualOperation(double x, double y, double width, double height) =>
        AddManualOperation(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddOffPageReference(double x, double y, double width, double height) =>
        AddOffPageReference(x, y, width, height, text: null, unit: DefaultUnit);

    /// <summary>Creates an unlabelled shape using the page default measurement unit.</summary>
    public VisioShape AddTriangle(double x, double y, double width, double height) =>
        AddTriangle(x, y, width, height, text: null, unit: DefaultUnit);

}
