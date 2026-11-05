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
[Cmdlet(VerbsData.Save, "OfficeAccess", SupportsShouldProcess = true, DefaultParameterSetName = "Document")]
[OutputType(typeof(AccessDocument))]
public sealed class SaveOfficeAccessCommand : PSCmdlet {
    /// <summary>Live database document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public AccessDocument Document { get; set; } = null!;

    /// <summary>Destination filename.</summary>
    [Parameter(Position = 0)]
    public string? Path { get; set; }

    /// <summary>Replace an existing destination through the canonical atomic writer.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Explicitly accept reported native-save loss.</summary>
    [Parameter] public SwitchParameter AllowLoss { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var path = Path == null ? null : OfficeFileCommandUtilities.ResolvePath(this, Path);
        if (!ShouldProcess(path ?? "Associated Access destination", "Save native Access database")) return;
        var options = new AccessSaveOptions {
            AllowLoss = AllowLoss,
            FileConflictPolicy = Force.IsPresent ? OfficeConversionFileConflictPolicy.Replace : OfficeConversionFileConflictPolicy.FailIfExists
        };
        var report = path == null ? Document.AssessSave(options) : Document.AssessSave(path, options);
        if (!AllowLoss) report.ThrowIfLoss();
        if (path == null) Document.Save(options); else Document.Save(path, options);
        if (PassThru.IsPresent) WriteObject(Document);
    }
}
