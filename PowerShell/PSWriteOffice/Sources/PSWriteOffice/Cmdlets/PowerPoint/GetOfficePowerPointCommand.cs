using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Loads an existing PowerPoint presentation.</summary>
/// <para>Returns an OfficeIMO <see cref="PowerPointPresentation"/> for downstream slide operations.</para>
/// <example>
///   <summary>Open a deck for editing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = Get-OfficePowerPoint -Path .\Quarterly.pptx</code>
///   <para>Reads <c>Quarterly.pptx</c> and exposes the presentation object.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPoint")]
public class GetOfficePowerPointCommand : PSCmdlet {
    /// <summary>Path to the .pptx file.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to open an encrypted presentation package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            var presentation = PowerPointDocumentService.LoadPresentation(resolvedPath, Password);
            WriteObject(presentation);
        } catch (FileNotFoundException ex) {
            WriteError(new ErrorRecord(ex, "FileNotFound", ErrorCategory.ObjectNotFound, Path));
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointLoadFailed", ErrorCategory.InvalidOperation, Path));
        }
    }
}