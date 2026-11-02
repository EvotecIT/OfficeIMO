using System.Management.Automation;
using OfficeIMO.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds a diamond shape to the current Visio page.</summary>
/// <example>
///   <summary>Add a decision shape.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeVisio -Path .\Flow.vsdx {
///     VisioDiamond -Key review -Text 'Approved?' -X 4in -Y 4in -Width 1.2in -Height 1in -FillColor '#FEF3C7'
/// }</code>
///   <para>Adds a decision diamond to the active Visio page.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioDiamond", DefaultParameterSetName = "Page")]
[Alias("VisioDiamond")]
[OutputType(typeof(VisioShape))]
public sealed class AddOfficeVisioDiamondCommand : OfficeVisioShapeCmdlet {
    /// <summary>Retains the native diamond dimensions.</summary>
    public AddOfficeVisioDiamondCommand() { Width = OfficeIMO.OfficeLength.Inches(1); }

    /// <summary>Text placed inside the shape.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <inheritdoc />
    protected override VisioShape CreateShape(VisioPage page) =>
        page.AddDiamond(X.ToInches(), Y.ToInches(), Width.ToInches(), Height.ToInches(), Text, VisioMeasurementUnit.Inches);
}
