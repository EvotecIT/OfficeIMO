using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Visio;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds a page to a Visio document and optionally executes nested DSL content.</summary>
/// <example>
///   <summary>Add a second diagram page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeVisio -Path .\Workbook.vsdx {
///     VisioPage -Name 'Architecture' -Width 11in -Height 8.5in {
///         VisioRectangle -Key api -Text 'API' -X 2in -Y 4in
///     }
/// }</code>
///   <para>Adds a named page and executes the nested shape DSL inside that page scope.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioPage", DefaultParameterSetName = "Document")]
[Alias("VisioPage")]
[OutputType(typeof(VisioPage))]
public sealed class AddOfficeVisioPageCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Target Visio document. Optional inside <c>New-OfficeVisio</c>.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioDocument? Document { get; set; }

    /// <summary>Page name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Page width.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Millimeters(210);

    /// <summary>Page height.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Millimeters(297);

    /// <summary>Measurement unit displayed by Visio, independent of geometry input units.</summary>
    [Parameter]
    public VisioMeasurementUnit DisplayUnit { get; set; } = VisioMeasurementUnit.Inches;

    /// <summary>Nested DSL content executed within this page scope.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = VisioDslContext.Current;
        var document = Document ?? (context ?? VisioDslContext.Require(this)).Document;
        var page = document.AddPage(Name, Width.ToInches(), Height.ToInches());
        page.DefaultUnit = DisplayUnit;
        page.ScaleMeasurementUnit = DisplayUnit;

        if (Content != null) {
            if (context != null) {
                using (context.Push(page)) {
                    Content.InvokeReturnAsIs();
                }
            } else {
                using (var scoped = VisioDslContext.Enter(document))
                using (scoped.Push(page)) {
                    Content.InvokeReturnAsIs();
                }
            }
        }

        WritePassThru(page);
    }
}
