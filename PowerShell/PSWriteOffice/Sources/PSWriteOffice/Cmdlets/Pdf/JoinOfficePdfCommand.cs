using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Joins multiple PDF files into a single PDF.</summary>
/// <example>
///   <summary>Join two PDFs in order.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$cover = '.\Examples\Documents\Cover.pdf'
/// $report = '.\Examples\Documents\Report.pdf'
/// Join-OfficePdf -Path $cover, $report -OutputPath .\Examples\Documents\Combined.pdf -PassThru
/// Get-OfficePdfInfo -Path .\Examples\Documents\Combined.pdf | Select-Object PageCount</code>
///   <para>Writes a single PDF containing the input documents in the requested order, then checks the result.</para>
/// </example>
/// <example>
///   <summary>Join encrypted sources and inspect the security decisions.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$result = Join-OfficePdf -Path .\Restricted.pdf, .\Appendix.pdf `
///     -Password 'source-password', $null -IgnorePermissionRestrictions `
///     -OutputPath .\Combined.pdf -PassThruReport
/// $result.Sources | Select-Object SourceIndex, PasswordAuthenticationRole, PermissionRestrictionsIgnored
/// $result.Decisions | Select-Object Structure, Mode, Action</code>
///   <para>A valid password remains mandatory. The explicit switch ignores usage flags after authentication and the report records every source decision.</para>
/// </example>
[Cmdlet(VerbsCommon.Join, "OfficePdf", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(PdfMergeReport))]
public sealed class JoinOfficePdfCommand : PSCmdlet
{
    /// <summary>Input PDF paths in output order.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string[] Path { get; set; } = System.Array.Empty<string>();

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>
    /// Passwords used to authenticate encrypted sources. Supply one value for every source, or one value to reuse for all sources.
    /// </summary>
    [Parameter]
    [AllowNull]
    public string?[]? Password { get; set; }

    /// <summary>
    /// After successful password authentication, explicitly ignore owner-imposed usage restrictions such as copying or assembly.
    /// This does not discover, bypass, or crack a missing password.
    /// </summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Emit the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Emit the per-source merge inventory, permission decisions, and output security readback.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <summary>Flatten visual annotation appearances before merging.</summary>
    [Parameter]
    public SwitchParameter FlattenVisualAnnotations { get; set; }

    /// <summary>Resize each merged page to a known OfficeIMO page size such as A4, Letter, or Custom.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficePdfPageSizeArgumentCompleter))]
    public string? PageSize { get; set; }

    /// <summary>Custom output page width in points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Custom output page height in points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Use the landscape orientation of the selected output page size.</summary>
    [Parameter]
    public SwitchParameter Landscape { get; set; }

    /// <summary>How source page content is fitted into the resized output page.</summary>
    [Parameter]
    public PdfPageResizeMode ResizeMode { get; set; } = PdfPageResizeMode.Fit;

    /// <summary>Margin, in points, reserved around resized page content.</summary>
    [Parameter]
    public double? ResizeMargin { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Path.Length == 0)
        {
            throw new PSArgumentException("Provide at least one input PDF path.", nameof(Path));
        }

        var passwords = ResolvePasswords(Path.Length);
        var resizeOptions = PdfCommandUtilities.CreatePageResizeOptions(
            PageSize,
            Width,
            Height,
            Landscape.IsPresent,
            ResizeMode,
            ResizeMargin,
            MyInvocation.BoundParameters.ContainsKey(nameof(PageSize)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Height)) ||
            Landscape.IsPresent ||
            MyInvocation.BoundParameters.ContainsKey(nameof(ResizeMode)) ||
            ResizeMargin.HasValue);

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write joined PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        var documents = Path
            .Select((path, index) => PdfDocument.Load(
                PdfCommandUtilities.ResolvePath(this, path),
                PdfCommandUtilities.CreateReadOptions(passwords[index], IgnorePermissionRestrictions.IsPresent)))
            .ToArray();
        var mergeOptions = new PdfMergeOptions
        {
            FlattenVisualAnnotations = FlattenVisualAnnotations.IsPresent,
            ResizePages = resizeOptions
        };
        var result = PdfDocument.MergeResult(mergeOptions, documents);
        result.ToDocument().Save(outputPath).RequireSuccess();

        if (PassThruReport.IsPresent)
        {
            WriteObject(result.Report);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private string?[] ResolvePasswords(int sourceCount)
    {
        if (Password is not { Length: > 0 })
        {
            return new string?[sourceCount];
        }

        if (Password.Length == 1)
        {
            return Enumerable.Repeat(Password[0], sourceCount).ToArray();
        }

        if (Password.Length != sourceCount)
        {
            throw new PSArgumentException(
                "Provide either one -Password value or one value for every -Path source.",
                nameof(Password));
        }

        return Password;
    }
}
