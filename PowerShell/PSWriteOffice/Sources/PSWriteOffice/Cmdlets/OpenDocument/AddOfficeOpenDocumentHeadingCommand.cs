using System.Management.Automation;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Services.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Adds a heading to an OpenDocument text document.</summary>
/// <example>
///   <summary>Add a level-two heading inside an OpenDocument DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeOpenDocumentHeading -Text 'Results' -Level 2</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeOpenDocumentHeading")]
[OutputType(typeof(OdtParagraph))]
public sealed class AddOfficeOpenDocumentHeadingCommand : PSCmdlet {
    /// <summary>OpenDocument text document. Omit inside New-OfficeOpenDocument -Content.</summary>
    [Parameter(ValueFromPipeline = true)]
    public OdtDocument? Document { get; set; }

    /// <summary>Heading text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Heading level from 1 through 10.</summary>
    [Parameter]
    [ValidateRange(1, 10)]
    public int Level { get; set; } = 1;

    /// <summary>Emit the created heading paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OdtDocument document = Document ?? OpenDocumentDslContext.Require(this).RequireDocument<OdtDocument>(this, "text");
        OdtParagraph heading = document.AddHeading(Text, Level);
        if (PassThru.IsPresent) WriteObject(heading);
    }
}
