using System.IO;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Visio;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Creates a Visio document with an initial page, returning it live or saving a completed content block.</summary>
/// <para>Without a block no file is written. A block publishes only after success; an existing destination requires Force. Geometry carries its own unit; DisplayUnit controls the page rulers.</para>
/// <example><summary>Create and save a service map.</summary>
/// <code>New-OfficeVisio -Path ./ServiceMap.vsdx -Title 'Service map' {
///     VisioRectangle -Key web -Text 'Web' -X 1in -Y 4in
///     VisioRectangle -Key api -Text 'API' -X 4in -Y 4in
///     VisioConnector -From web -To api -Label 'calls'
/// }</code></example>
/// <example><summary>Create a live landscape diagram and save it explicitly.</summary>
/// <code>$diagram = New-OfficeVisio -Width 11in -Height 8.5in
/// Save-OfficeVisio -Document $diagram -Path ./Landscape.vsdx</code></example>
[Cmdlet(VerbsCommon.New, "OfficeVisio", SupportsShouldProcess = true)]
[OutputType(typeof(VisioDocument), typeof(FileInfo))]
public sealed class NewOfficeVisioCommand : OfficeLifecycleCmdlet<VisioDocument> {
    /// <summary>Optional associated destination; required with a content block.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Initial page name.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string PageName { get; set; } = "Page-1";

    /// <summary>Initial page width; defaults to A4 portrait.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Millimeters(210);

    /// <summary>Initial page height; defaults to A4 portrait.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Millimeters(297);

    /// <summary>Measurement unit displayed by Visio, independent of geometry input units.</summary>
    [Parameter]
    public VisioMeasurementUnit DisplayUnit { get; set; } = VisioMeasurementUnit.Inches;

    /// <summary>Optional document title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional document author.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Ask Visio to recalculate layout and connector routing on opening.</summary>
    [Parameter]
    public SwitchParameter RequestRecalcOnOpen { get; set; }

    /// <summary>Use Visio masters for supported built-in stencil shapes when saving.</summary>
    [Parameter]
    public SwitchParameter UseMastersByDefault { get; set; }

    /// <summary>Allow a completed content block to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => RunDocument(Path, isNew: true, force: Force.IsPresent,
        path => {
            VisioDocument document = path == null ? VisioDocument.Create() : VisioDocument.Create(path);
            document.Title = Title;
            document.Author = Author;
            document.UseMastersByDefault = UseMastersByDefault.IsPresent;
            if (RequestRecalcOnOpen) document.RequestRecalcOnOpen();
            VisioPage page = document.AddPage(PageName, Width.ToInches(), Height.ToInches());
            page.DefaultUnit = DisplayUnit;
            page.ScaleMeasurementUnit = DisplayUnit;
            return document;
        },
        (document, path, allowLoss) => { document.Save(path); return null; });

    /// <inheritdoc />
    protected override void InvokeContent(VisioDocument document) {
        using var context = VisioDslContext.Enter(document);
        using (context.Push(document.Pages[0])) Content!.InvokeReturnAsIs();
    }
}
