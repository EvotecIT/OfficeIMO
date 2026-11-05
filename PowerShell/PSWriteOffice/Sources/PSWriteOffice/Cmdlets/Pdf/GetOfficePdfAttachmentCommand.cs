using System;
using System.IO;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Lists embedded files, or extracts them with -OutputDirectory.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfAttachment", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(PdfExtractedAttachment), typeof(FileInfo))]
public sealed class GetOfficePdfAttachmentCommand : PSCmdlet {
    /// <summary>Path for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Name for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public string? Name { get; set; }
    /// <summary>OutputDirectory for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public string? OutputDirectory { get; set; }
    /// <summary>NoClobber for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter NoClobber { get; set; }
    /// <summary>Password for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>IgnorePermissionRestrictions for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>LoadOptions for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>LiteralPath for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Document for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => PdfInspectionSource.Visit(this, ParameterSetName, Path, LiteralPath, Document, LoadOptions, Password, IgnorePermissionRestrictions, Inspect);

    private void Inspect(PdfDocument document) {
        string? directory = OutputDirectory == null ? null : OfficePath.ResolveOutput(SessionState, OutputDirectory, directory: true);
        if (directory != null && !ShouldProcess(directory, "Extract PDF attachments")) return;
        foreach (var attachment in document.Attachments.Extract()) {
            if (Name != null && !string.Equals(attachment.Name, Name, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(attachment.FileName, Name, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(attachment.UnicodeFileName, Name, StringComparison.OrdinalIgnoreCase)) continue;
            if (directory == null) { WriteObject(attachment); continue; }
            var name = string.IsNullOrWhiteSpace(attachment.UnicodeFileName) ? attachment.FileName : attachment.UnicodeFileName!;
            WriteObject(new FileInfo(WriteExtraction(directory, name, attachment.Bytes)));
        }
    }

    private string WriteExtraction(string directory, string name, byte[] bytes) {
        if (!NoClobber) return OfficeFileOutput.WriteUnique(directory, name, bytes);
        string path = System.IO.Path.Combine(directory, OfficeFileOutput.GetSafeFileName(name));
        OfficeFileOutput.Write(path, bytes, overwrite: false);
        return path;
    }
}
