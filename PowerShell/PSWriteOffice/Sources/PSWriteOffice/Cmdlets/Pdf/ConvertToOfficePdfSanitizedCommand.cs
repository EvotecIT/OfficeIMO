using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Removes or quarantines active PDF content and embedded payloads with post-save proof.</summary>
/// <example>
///   <summary>Sanitize a PDF and inspect what was removed.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$result = ConvertTo-OfficePdfSanitized -Path .\Input.pdf -OutputPath .\Safe.pdf</code>
///   <para>Writes the proven full-rewrite result and returns findings, mutation plan, and quarantine data.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfSanitized", SupportsShouldProcess = true)]
[OutputType(typeof(PdfSanitizationResult))]
public sealed class ConvertToOfficePdfSanitizedCommand : PSCmdlet
{
    /// <summary>Source PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Allowed actions, URI schemes, embedded-file policy, and rich-media policy.</summary>
    [Parameter]
    public PdfSanitizationOptions? Options { get; set; }

    /// <summary>Optional bounded PDF parsing and password settings.</summary>
    [Parameter]
    public PdfLoadOptions? ReadOptions { get; set; }

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed modification restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, "Sanitize PDF active content and embedded payloads")) return;
        var result = PdfCommandUtilities.LoadDocument(
            input,
            PdfCommandUtilities.CreateReadOptions(
                ReadOptions,
                Password,
                IgnorePermissionRestrictions.IsPresent)).Sanitize(Options);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        File.WriteAllBytes(output, result.ToBytes());
        WriteObject(result);
    }
}
