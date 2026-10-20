using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;
using OfficeIMO.Core.Internal;
namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Writes a validated EPUB package with explicit conflict and loss policies.</summary>
/// <example>
///   <summary>Writes a validated EPUB package with explicit conflict and loss policies.</summary>
///   <code>$book | Save-OfficeEpub -Path ./operations.epub</code>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeEpub", SupportsShouldProcess = true)]
[OutputType(typeof(EpubPublication))]
public sealed class SaveOfficeEpubCommand : PSCmdlet {
    /// <summary>Live editable publication.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public EpubPublication Publication { get; set; } = null!;

    /// <summary>Destination EPUB path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Permit reductions reported by the package writer.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Explicitly remove signatures invalidated by package edits. Requires AllowLoss if reported as loss.</summary>
    [Parameter]
    public SwitchParameter RemoveInvalidatedSignatures { get; set; }

    /// <summary>Maximum compressed output size.</summary>
    [Parameter]
    [ValidateRange(1L, int.MaxValue)]
    public long MaxOutputBytes { get; set; } = 134217728;

    /// <summary>Variable receiving the writer preservation report.</summary>
    [Parameter]
    public string? ReportVariable { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var path = OfficeFileCommandUtilities.ResolvePath(this, Path);
        if (!ShouldProcess(path, "Save EPUB publication")) return;
        var result = Publication.Write(new EpubWriteOptions {
            RemoveInvalidatedSignatures = RemoveInvalidatedSignatures.IsPresent, MaxOutputBytes = MaxOutputBytes
        });
        if (!AllowLoss.IsPresent) result.Report.RequireNoLoss();
        OfficeFileOutput.Write(path, result.Bytes, Force.IsPresent);
        if (!string.IsNullOrWhiteSpace(ReportVariable)) SessionState.PSVariable.Set(ReportVariable, result.Report);
        if (PassThru.IsPresent) WriteObject(Publication);
    }
}
