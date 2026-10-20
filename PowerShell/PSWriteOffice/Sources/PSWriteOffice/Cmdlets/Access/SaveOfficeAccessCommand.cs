using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Writes a qualified native database with strict loss protection.</summary>
/// <example>
///   <summary>Writes a qualified native database with strict loss protection.</summary>
///   <code>$database | Save-OfficeAccess -Path ./contacts.accdb</code>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeAccess", SupportsShouldProcess = true)]
[OutputType(typeof(AccessDocument))]
public sealed class SaveOfficeAccessCommand : PSCmdlet {
    /// <summary>Live database document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public AccessDocument Document { get; set; } = null!;

    /// <summary>Destination filename.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Replace an existing destination through the canonical atomic writer.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var path = OfficeFileCommandUtilities.ResolvePath(this, Path);
        if (!ShouldProcess(path, "Save native Access database")) return;
        var options = new AccessSaveOptions {
            FileConflictPolicy = Force.IsPresent ? OfficeConversionFileConflictPolicy.Replace : OfficeConversionFileConflictPolicy.FailIfExists
        };
        Document.AssessSave(path, options).RequireNoLoss();
        Document.Save(path, options);
        if (PassThru.IsPresent) WriteObject(Document);
    }
}
