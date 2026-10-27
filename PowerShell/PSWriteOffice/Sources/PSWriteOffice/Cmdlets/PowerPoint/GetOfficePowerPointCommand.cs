using PSWriteOffice.Services;
using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Loads an existing PowerPoint presentation.</summary>
/// <para>Returns an OfficeIMO <see cref="PowerPointDocument"/> for downstream slide operations.</para>
/// <example>
///   <summary>Open a deck for editing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = Get-OfficePowerPoint -Path .\Quarterly.pptx</code>
///   <para>Reads <c>Quarterly.pptx</c> and exposes the presentation object.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPoint")]
public class GetOfficePowerPointCommand : PSCmdlet {
    /// <summary>Path to the .pptx file.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ValueFromPipeline = true)]
    [ValidateNotNullOrEmpty]
    [Alias("FilePath", "FullName", "LiteralPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to open an encrypted presentation package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var resolvedPath = OfficeFileCommandUtilities.ResolvePath(this, Path);
            var presentation = PowerPointDocumentService.LoadDocument(resolvedPath, Password);
            WriteObject(presentation);
        } catch (FileNotFoundException ex) {
            WriteError(new ErrorRecord(ex, global::OfficeIMO.OfficeDiagnosticCodes.FileNotFound, ErrorCategory.ObjectNotFound, Path));
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointLoadFailed", ErrorCategory.InvalidOperation, Path));
        }
    }
}