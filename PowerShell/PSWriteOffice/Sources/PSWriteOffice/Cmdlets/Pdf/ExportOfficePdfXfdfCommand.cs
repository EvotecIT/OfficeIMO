using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Exports readable PDF form field values as XFDF.</summary>
[Cmdlet(VerbsData.Export, "OfficePdfXfdf", SupportsShouldProcess = true)]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ExportOfficePdfXfdfCommand : PSCmdlet
{
    /// <summary>Source PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional XFDF output path. Without it, the command returns XML.</summary>
    [Parameter(Position = 1)]
    public string? OutputPath { get; set; }

    /// <summary>Optional bounded PDF parsing and password settings.</summary>
    [Parameter]
    public PdfLoadOptions? ReadOptions { get; set; }

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed form-reading restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Return the written file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var readOptions = PdfCommandUtilities.CreateReadOptions(
            ReadOptions,
            Password,
            IgnorePermissionRestrictions.IsPresent);
        var xfdf = PdfCommandUtilities.LoadDocument(input, readOptions).Forms.ExportXfdf();
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(xfdf);
            return;
        }
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath!);
        if (!ShouldProcess(output, "Write PDF form data as XFDF")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        File.WriteAllText(output, xfdf, new UTF8Encoding(false));
        if (PassThru.IsPresent) WriteObject(new FileInfo(output));
    }
}
