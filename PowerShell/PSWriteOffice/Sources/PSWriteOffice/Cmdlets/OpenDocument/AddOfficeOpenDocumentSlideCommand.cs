using System.Management.Automation;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Services.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Adds a slide to an OpenDocument presentation and optionally runs nested slide content.</summary>
/// <example>
///   <summary>Add a slide with positioned text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeOpenDocumentSlide -Name 'Summary' -Content {
///     Add-OfficeOpenDocumentTextBox -Text 'Quarterly summary' -X 2 -Y 2 -Width 20 -Height 3
/// }</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeOpenDocumentSlide")]
[OutputType(typeof(OdpSlide))]
public sealed class AddOfficeOpenDocumentSlideCommand : PSCmdlet {
    /// <summary>OpenDocument presentation. Omit inside New-OfficeOpenDocument -Content.</summary>
    [Parameter(ValueFromPipeline = true)]
    public OdpPresentation? Document { get; set; }

    /// <summary>Optional unique slide name.</summary>
    [Parameter(Position = 0)]
    public string? Name { get; set; }

    /// <summary>Nested slide commands that use this slide as their current target.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit the created slide.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OpenDocumentDslContext? context = OpenDocumentDslContext.Current;
        OdpPresentation document = Document ?? OpenDocumentDslContext.Require(this).RequireDocument<OdpPresentation>(this, "presentation");
        OdpSlide slide = document.AddSlide(Name);
        if (Content != null) {
            if (context == null) throw new PSInvalidOperationException("Nested -Content requires an active New-OfficeOpenDocument -Content scope. For object composition, pass the returned slide to Add-OfficeOpenDocumentTextBox.");
            using (context.Push(slide)) Content.InvokeReturnAsIs();
        }
        if (PassThru.IsPresent) WriteObject(slide);
    }
}
