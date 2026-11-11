using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Writes a smaller or sanitized copy of a PDF.</summary>
/// <example><summary>Use the native OfficePdf command.</summary><code>Optimize-OfficePdf -Path .\Reports\Annual-Report.pdf -OutputPath .\Reports\Annual-Report-Small.pdf</code></example>
[Cmdlet("Optimize", "OfficePdf", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(object))]
public sealed class OptimizeOfficePdfCommand : OfficeCmdlet {
    private readonly List<string> _inputs = new();
    /// <summary>Path for this operation.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathSanitize", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathSanitizeDirectory", Mandatory = true, Position = 0)]
    public string[]? Path { get; set; }
    /// <summary>OutputPath for this operation.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true)]
    [Parameter(ParameterSetName = "PathSanitize", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathSanitize", Mandatory = true)]
    public string? OutputPath { get; set; }
    /// <summary>NoCompressStream for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter NoCompressStream { get; set; }
    /// <summary>KeepUnreferencedObject for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter KeepUnreferencedObject { get; set; }
    /// <summary>KeepDuplicateStream for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter KeepDuplicateStream { get; set; }
    /// <summary>AllowLarger for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter AllowLarger { get; set; }
    /// <summary>MinimumStreamCompressionSize for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public int? MinimumStreamCompressionSize { get; set; }
    /// <summary>Argument transformation also accepts a plain string (with a warning) or a PSCredential (its password).</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathSanitize")]
    [Parameter(ParameterSetName = "PathSanitizeDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathSanitize")]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory")]
    [OfficePasswordTransformation]
    public System.Security.SecureString? Password { get; set; }
    /// <summary>After successful authentication only; never bypasses a missing password.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathSanitize")]
    [Parameter(ParameterSetName = "PathSanitizeDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathSanitize")]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>LoadOptions for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathSanitize")]
    [Parameter(ParameterSetName = "PathSanitizeDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathSanitize")]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>NoClobber for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathSanitize")]
    [Parameter(ParameterSetName = "PathSanitizeDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathSanitize")]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory")]
    public SwitchParameter NoClobber { get; set; }
    /// <summary>PassThru for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "PathSanitize")]
    [Parameter(ParameterSetName = "PathSanitizeDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "LiteralPathSanitize")]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>OutputDirectory for this operation.</summary>
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PathSanitizeDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory", Mandatory = true)]
    public string? OutputDirectory { get; set; }
    /// <summary>Sanitize for this operation.</summary>
    [Parameter(ParameterSetName = "PathSanitize", Mandatory = true)]
    [Parameter(ParameterSetName = "PathSanitizeDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathSanitize", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory", Mandatory = true)]
    public SwitchParameter Sanitize { get; set; }
    /// <summary>Hidden-data categories that -Sanitize removes, as OfficeHiddenDataOptions or a hashtable such as @{ Kinds = 'ActiveContent, Attachments, DocumentProperties' }. Defaults to attachments and active content: scripts, unsafe actions and links, and rich media.</summary>
    [Parameter(ParameterSetName = "PathSanitize")]
    [Parameter(ParameterSetName = "PathSanitizeDirectory")]
    [Parameter(ParameterSetName = "LiteralPathSanitize")]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory")]
    [Alias("SanitizationOptions")]
    [OfficeOptionsTransformation(typeof(OfficeHiddenDataOptions))]
    public OfficeHiddenDataOptions? HiddenDataOptions { get; set; }
    /// <summary>LiteralPath for this operation.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathSanitize", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathSanitizeDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[]? LiteralPath { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        _inputs.AddRange(OfficePath.ResolveInputs(this, LiteralPath ?? Path ?? Array.Empty<string>(), LiteralPath != null));
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        foreach (var pair in PdfFileCommandInputs.Resolve(this, null, _inputs.ToArray(), OutputPath, OutputDirectory, NoClobber)) {
            if (!ShouldProcess(pair.Output, "Optimize PDF")) continue;
            using var document = PdfDocument.Load(pair.Input, PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions));
            if (Sanitize) {
                var result = document.RemoveHiddenData(HiddenDataOptions ?? new OfficeHiddenDataOptions {
                    Kinds = OfficeHiddenDataKind.ActiveContent | OfficeHiddenDataKind.Attachments
                });
                PdfCommandUtilities.EnsureDirectory(pair.Output);
                OfficeIMO.Core.Internal.OfficeFileOutput.Write(pair.Output, !NoClobber, temporary => File.WriteAllBytes(temporary, result.ToArray()));
                WriteObject(PassThru ? (object)result : new FileInfo(pair.Output));
            } else {
                var result = document.Optimization.Apply(new PdfOptimizationOptions {
                    CompressUnfilteredStreams = !NoCompressStream,
                    RemoveUnreferencedObjects = !KeepUnreferencedObject,
                    DeduplicateIdenticalStreams = !KeepDuplicateStream,
                    KeepOriginalWhenNotSmaller = !AllowLarger,
                    MinimumStreamCompressionBytes = MinimumStreamCompressionSize ?? 128
                });
                PdfCommandUtilities.EnsureDirectory(pair.Output);
                using (var output = result.ToDocument()) {
                    output.Save(pair.Output, new PdfSaveOptions { ConflictPolicy = NoClobber
                        ? OfficeConversionFileConflictPolicy.FailIfExists : OfficeConversionFileConflictPolicy.Replace });
                }
                WriteObject(PassThru ? (object)result : new FileInfo(pair.Output));
            }
        }
    }
}
