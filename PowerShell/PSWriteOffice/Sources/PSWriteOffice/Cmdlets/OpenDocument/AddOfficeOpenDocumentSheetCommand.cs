using System.Management.Automation;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Services.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Adds a worksheet to an OpenDocument spreadsheet and optionally runs nested cell content.</summary>
/// <example>
///   <summary>Add a worksheet inside an OpenDocument DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeOpenDocumentSheet -Name 'Data' -Content {
///     Set-OfficeOpenDocumentCell -Row 0 -Column 0 -Value 'Status'
/// }</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeOpenDocumentSheet")]
[OutputType(typeof(OdsSheet))]
public sealed class AddOfficeOpenDocumentSheetCommand : PSCmdlet {
    /// <summary>OpenDocument spreadsheet. Omit inside New-OfficeOpenDocument -Content.</summary>
    [Parameter(ValueFromPipeline = true)]
    public OdsDocument? Document { get; set; }

    /// <summary>Worksheet name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Nested cell commands that use this worksheet as their current target.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit the created worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OpenDocumentDslContext? context = OpenDocumentDslContext.Current;
        OdsDocument document = Document ?? OpenDocumentDslContext.Require(this).RequireDocument<OdsDocument>(this, "spreadsheet");
        OdsSheet sheet = document.AddSheet(Name);
        if (Content != null) {
            if (context == null) throw new PSInvalidOperationException("Nested -Content requires an active New-OfficeOpenDocument -Content scope. For object composition, pass the returned sheet to Set-OfficeOpenDocumentCell.");
            using (context.Push(sheet)) Content.InvokeReturnAsIs();
        }
        if (PassThru.IsPresent) WriteObject(sheet);
    }
}
