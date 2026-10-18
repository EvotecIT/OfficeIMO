using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Adds a page to a Visio document and optionally executes nested DSL content.</summary>
/// <example>
///   <summary>Add a second diagram page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\Workbook.vsdx {
///     VisioPage -Name 'Architecture' {
///         VisioRectangle -Key api -Text 'API' -X 2 -Y 4
///     }
/// }</code>
///   <para>Adds a named page and executes the nested shape DSL inside that page scope.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeVisioPage")]
[Alias("VisioPage")]
[OutputType(typeof(VisioPage))]
public sealed class AddOfficeVisioPageCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Target Visio document. Optional inside <c>OfficeVisio</c>.</summary>
    [Parameter(ValueFromPipeline = true)]
    public VisioDocument? Document { get; set; }

    /// <summary>Page name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Page width.</summary>
    [Parameter]
    public double Width { get; set; } = 8.26771653543307;

    /// <summary>Page height.</summary>
    [Parameter]
    public double Height { get; set; } = 11.69291338582677;

    /// <summary>Measurement unit for width and height.</summary>
    [Parameter]
    public VisioMeasurementUnit Unit { get; set; } = VisioMeasurementUnit.Inches;

    /// <summary>Nested DSL content executed within this page scope.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = VisioDslContext.Current;
        var document = Document ?? (context ?? VisioDslContext.Require(this)).Document;
        var page = document.AddPage(Name, Width, Height, Unit);

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
