using System.IO;
using System.Management.Automation;
using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Saves a native EPUB publication, rejecting loss unless explicitly allowed.</summary>
/// <para>Save replaces its destination after validation. PassThru returns the native write report.</para>
/// <example><summary>Save a live publication to its associated path and inspect the report.</summary>
/// <code>$report = $book | Save-OfficeEpub -PassThru</code></example>
[Cmdlet(VerbsData.Save, "OfficeEpub", SupportsShouldProcess = true, DefaultParameterSetName = "Document")]
[OutputType(typeof(EpubWriteReport))]
public sealed class SaveOfficeEpubCommand : PSCmdlet {
    /// <summary>Live editable publication.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public EpubPublication Document { get; set; } = null!;

    /// <summary>Destination path; omission uses the native publication's associated path.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Permit reductions reported by the package writer.</summary>
    [Parameter]
    public SwitchParameter AllowLoss { get; set; }

    /// <summary>Explicitly remove signatures invalidated by edits. Reported loss still requires AllowLoss.</summary>
    [Parameter]
    public SwitchParameter RemoveInvalidatedSignatures { get; set; }

    /// <summary>Maximum compressed output size.</summary>
    [Parameter, ValidateRange(1L, int.MaxValue)]
    public long MaxOutputBytes { get; set; } = 134217728;

    /// <summary>Return the native write report.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string destination = Path ?? Document.FilePath ?? throw new PSArgumentException("Supply Path when the publication has no associated save path.", nameof(Path));
        string path = OfficeFileCommandUtilities.ResolvePath(this, destination);
        if (!ShouldProcess(path, "Save EPUB publication")) return;
        EpubWriteResult result = Document.Write(new EpubWriteOptions {
            RemoveInvalidatedSignatures = RemoveInvalidatedSignatures.IsPresent,
            MaxOutputBytes = MaxOutputBytes
        });
        if (!AllowLoss) result.Report.RequireNoLoss();
        OfficeFileOutput.Write(path, result.Bytes, true);
        Document.FilePath = path;
        if (PassThru) WriteObject(result.Report);
    }
}
