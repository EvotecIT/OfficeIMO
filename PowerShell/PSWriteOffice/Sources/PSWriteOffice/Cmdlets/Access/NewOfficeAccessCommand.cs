using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Creates a native Access database without a provider, seed file or execution.</summary>
/// <para>A path associates a live model without writing it; a content block atomically saves and closes it. The .mdb or .accdb extension selects the native file family.</para>
/// <example><summary>Create a contacts table.</summary><code>New-OfficeAccess -Path ./Contacts.accdb {
///     AccessTable Contacts { AccessColumn ContactId -Type AutoNumber; AccessColumn DisplayName -MaxLength 100 }
/// }</code></example>
[Cmdlet(VerbsCommon.New, "OfficeAccess", DefaultParameterSetName = "Live", SupportsShouldProcess = true)]
[OutputType(typeof(AccessDocument), typeof(FileInfo))]
public sealed class NewOfficeAccessCommand : OfficeLifecycleCmdlet<AccessDocument> {
    /// <summary>Optional live association; required for a completed content block.</summary>
    [Parameter(Position = 0, ParameterSetName = "Live")]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public string? Path { get; set; }
    /// <summary>Authoring block completed before atomic publication.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Content"), ValidateNotNull]
    public override ScriptBlock? Content { get => base.Content; set => base.Content = value; }
    /// <summary>Accept reductions reported by native save.</summary>
    [Parameter(ParameterSetName = "Content")]
    public override SwitchParameter AllowLoss { get => base.AllowLoss; set => base.AllowLoss = value; }
    /// <summary>Return the saved FileInfo after the block completes.</summary>
    [Parameter(ParameterSetName = "Content")]
    public override SwitchParameter PassThru { get => base.PassThru; set => base.PassThru = value; }
    /// <summary>Permit replacement of the associated destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }
    /// <summary>Database title.</summary>
    [Parameter]
    public string? Title { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        RunDocument(Path, true, Force, path => {
            var options = new AccessCreateOptions { DatabaseTitle = Title,
                Format = path != null && System.IO.Path.GetExtension(path).Equals(".mdb", StringComparison.OrdinalIgnoreCase) ? AccessFileFormat.Mdb : AccessFileFormat.Accdb };
            return path == null ? AccessDocument.Create(options) : AccessDocument.Create(path, options);
        }, (document, path, allowLoss) => {
            document.Save(path, new AccessSaveOptions { AllowLoss = allowLoss, FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace });
            return null;
        });
    }
    /// <inheritdoc />
    protected override void InvokeContent(AccessDocument document) {
        using (OfficeDocumentContext.Push(document)) Content!.InvokeReturnAsIs();
    }
}
