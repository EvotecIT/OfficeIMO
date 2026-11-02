using System.Management.Automation;
using OfficeIMO.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds an ellipse shape to the current Visio page.</summary>
/// <example>
///   <summary>Add a start node.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeVisio -Path .\Flow.vsdx {
///     VisioEllipse -Key start -Text 'Start' -X 1in -Y 5in -Width 1.4in -Height 0.7in -FillColor '#DCFCE7'
/// }</code>
///   <para>Adds an ellipse to the active Visio page.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioEllipse", DefaultParameterSetName = "Page")]
[Alias("VisioEllipse")]
[OutputType(typeof(VisioShape))]
public sealed class AddOfficeVisioEllipseCommand : OfficeVisioShapeCmdlet {
    /// <summary>Text placed inside the shape.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <inheritdoc />
    protected override VisioShape CreateShape(VisioPage page) =>
        page.AddEllipse(X.ToInches(), Y.ToInches(), Width.ToInches(), Height.ToInches(), Text, VisioMeasurementUnit.Inches);
}
