using System.IO;
using System.Management.Automation;
using OfficeIMO.Visio;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Creates a new OfficeIMO.Visio document with an initial page and optional DSL content.</summary>
/// <example>
///   <summary>Create a simple service map.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\ServiceMap.vsdx -Title 'Service map' -RequestRecalcOnOpen {
///     VisioRectangle -Key web -Text 'Web' -X 1 -Y 4 -FillColor LightBlue
///     VisioRectangle -Key api -Text 'API' -X 4 -Y 4 -FillColor LightGreen
///     VisioConnector -From web -To api -EndArrow Triangle -Label 'calls'
/// }</code>
///   <para>Creates an editable .vsdx diagram with two shapes and a connector.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeVisio", SupportsShouldProcess = true)]
[Alias("OfficeVisio", "VisioNew")]
[OutputType(typeof(VisioDocument), typeof(FileInfo))]
public sealed class NewOfficeVisioCommand : PSCmdlet {
    /// <summary>Destination .vsdx path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>DSL script block describing Visio pages, shapes, and connectors.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Initial page name.</summary>
    [Parameter]
    public string PageName { get; set; } = "Page-1";

    /// <summary>Initial page width.</summary>
    [Parameter]
    public double Width { get; set; } = 8.26771653543307;

    /// <summary>Initial page height.</summary>
    [Parameter]
    public double Height { get; set; } = 11.69291338582677;

    /// <summary>Measurement unit for page width and height.</summary>
    [Parameter]
    public VisioMeasurementUnit Unit { get; set; } = VisioMeasurementUnit.Inches;

    /// <summary>Optional document title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional document author.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Ask Visio to recalculate layout and connector routing when the document opens.</summary>
    [Parameter]
    public SwitchParameter RequestRecalcOnOpen { get; set; }

    /// <summary>Use Visio masters for supported built-in stencil shapes when saving.</summary>
    [Parameter]
    public SwitchParameter UseMastersByDefault { get; set; }

    /// <summary>Skip saving and emit the document object.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the document object instead of the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (NoSave.IsPresent && Open.IsPresent) {
            throw new PSArgumentException("-Open cannot be used with -NoSave because no file is written. Save the returned document explicitly, then use -Open on Save-OfficeVisio.", nameof(Open));
        }

        var fullPath = VisioCommandUtilities.ResolvePath(this, Path);
        if (!NoSave.IsPresent) {
            if (!ShouldProcess(fullPath, "Write new Visio document")) {
                return;
            }

            VisioCommandUtilities.EnsureDirectory(fullPath);
        }

        var document = VisioDocument.Create(fullPath);
        document.Title = Title;
        document.Author = Author;
        document.UseMastersByDefault = UseMastersByDefault.IsPresent;

        if (RequestRecalcOnOpen.IsPresent) {
            document.RequestRecalcOnOpen();
        }

        var page = document.AddPage(PageName, Width, Height, Unit);
        if (Content != null) {
            using (var context = VisioDslContext.Enter(document))
            using (context.Push(page)) {
                Content.InvokeReturnAsIs();
            }
        }

        if (NoSave.IsPresent) {
            WriteObject(document);
            return;
        }

        document.Save();

        if (Open.IsPresent) {
            FileOpenService.Open(fullPath);
        }

        if (PassThru.IsPresent) {
            WriteObject(new FileInfo(fullPath));
        }
    }
}
