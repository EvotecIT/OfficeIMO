using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Saves a project through the qualified native codec with loss protection.</summary>
/// <example>
///   <summary>Saves a project through the qualified native codec with loss protection.</summary>
///   <code>$project | Save-OfficeProject -Path ./delivery.mpp</code>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeProject", SupportsShouldProcess = true)]
[OutputType(typeof(ProjectDocument))]
public sealed class SaveOfficeProjectCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Destination filename; its extension selects the format.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Explicit generation or exchange format.</summary>
    [Parameter]
    public ProjectFileFormat Format { get; set; } = ProjectFileFormat.Automatic;

    /// <summary>Permit reported conversion loss. Unsupported codecs still fail.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Variable receiving the preservation and conversion assessment.</summary>
    [Parameter]
    public string? ReportVariable { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var path = OfficeFileCommandUtilities.ResolvePath(this, Path);
        if (!ShouldProcess(path, "Save project")) return;
        var options = new ProjectSaveOptions {
            Format = Format,
            LossPolicy = AllowLoss.IsPresent ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block,
            FileConflictPolicy = Force.IsPresent ? OfficeConversionFileConflictPolicy.Replace : OfficeConversionFileConflictPolicy.FailIfExists
        };
        var report = Document.AssessSave(path, options);
        if (!string.IsNullOrWhiteSpace(ReportVariable)) SessionState.PSVariable.Set(ReportVariable, report);
        report.ThrowIfErrors();
        Document.Save(path, options);
        if (PassThru.IsPresent) WriteObject(Document);
    }
}
