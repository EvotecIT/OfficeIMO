using System.Management.Automation;
using System.Security;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Lists image resources, or extracts them with -OutputDirectory.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(PdfExtractedImage), typeof(FileInfo))]
public sealed class GetOfficePdfImageCommand : PSCmdlet {
    /// <summary>Path for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Page for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [OfficePageRangeTransformation]
    public OfficePageRange? Page { get; set; }
    /// <summary>OutputDirectory for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public string? OutputDirectory { get; set; }
    /// <summary>BaseName for PDF inspection or extraction.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    public string BaseName { get; set; } = "image";
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
        if (directory != null && !ShouldProcess(directory, "Extract PDF images")) return;
        var images = Page == null ? document.Images.Extract() : document.Images.Extract(Page.ToString());
        int savedIndex = 0;
        foreach (var image in images) {
            if (directory == null) { WriteObject(image); continue; }
            if (!image.IsImageFile || string.IsNullOrWhiteSpace(image.FileExtension)) continue;
            string name = $"{BaseName}-{++savedIndex:D4}.{image.FileExtension!.TrimStart('.')}";
            WriteObject(new FileInfo(WriteExtraction(directory, name, image.Bytes)));
        }
    }

    private string WriteExtraction(string directory, string name, byte[] bytes) {
        if (!NoClobber) return OfficeFileOutput.WriteUnique(directory, name, bytes);
        string path = System.IO.Path.Combine(directory, OfficeFileOutput.GetSafeFileName(name));
        OfficeFileOutput.Write(path, bytes, overwrite: false);
        return path;
    }
}
