using System.Management.Automation;
using OfficeIMO.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds a rectangle shape to the current Visio page.</summary>
/// <example>
///   <summary>Add a labeled process box.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeVisio -Path .\Flow.vsdx {
///     VisioRectangle -Key intake -Text 'Intake' -X 1.5in -Y 4in -Width 1.7in -Height 0.8in -FillColor '#E0F2FE'
/// }</code>
///   <para>Adds a rectangle and registers a key for later connector commands.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioRectangle")]
[Alias("VisioRectangle")]
[OutputType(typeof(VisioShape))]
public sealed class AddOfficeVisioRectangleCommand : OfficeVisioShapeCmdlet {
    /// <summary>Text placed inside the shape.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <inheritdoc />
    protected override VisioShape CreateShape(VisioPage page) =>
        page.AddRectangle(X.ToInches(), Y.ToInches(), Width.ToInches(), Height.ToInches(), Text, VisioMeasurementUnit.Inches);
}
