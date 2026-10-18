using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Converts a PDF file to HTML through the first-party OfficeIMO HTML/PDF adapter.</summary>
/// <example>
///   <summary>Export a PDF as semantic HTML.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\report.pdf { Add-OfficePdfParagraph -Text 'Ready' }
/// ConvertTo-OfficePdfHtml -Path .\report.pdf -OutputPath .\report.html</code>
///   <para>Writes HTML generated from the OfficeIMO logical PDF read model.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfHtml", SupportsShouldProcess = true)]
[Alias("ConvertTo-PdfHtml")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficePdfHtmlCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional page ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Password used to read a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Optional output HTML file path.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>PDF to HTML profile to use when <see cref="Options"/> is not supplied.</summary>
    [Parameter]
    public PdfHtmlProfile Profile { get; set; } = PdfHtmlProfile.Semantic;

    /// <summary>Controls whether extracted images are embedded or represented as placeholders.</summary>
    [Parameter]
    public PdfHtmlImageExportMode ImageExportMode { get; set; } = PdfHtmlImageExportMode.EmbeddedDataUri;

    /// <summary>Maximum extracted image byte length that may be embedded into generated HTML. Set to 0 to disable embedding.</summary>
    [Parameter]
    public long? MaxEmbeddedImageBytes { get; set; }

    /// <summary>Do not emit PDF metadata into the generated HTML.</summary>
    [Parameter]
    public SwitchParameter NoMetadata { get; set; }

    /// <summary>Do not emit page wrapper elements.</summary>
    [Parameter]
    public SwitchParameter NoPageContainers { get; set; }

    /// <summary>Do not emit image placeholders.</summary>
    [Parameter]
    public SwitchParameter NoImagePlaceholders { get; set; }

    /// <summary>Include link annotation placeholders.</summary>
    [Parameter]
    public SwitchParameter IncludeLinkAnnotations { get; set; }

    /// <summary>Include AcroForm widget placeholders.</summary>
    [Parameter]
    public SwitchParameter IncludeFormWidgets { get; set; }

    /// <summary>Emit an HTML fragment instead of a complete document shell.</summary>
    [Parameter]
    public SwitchParameter Fragment { get; set; }

    /// <summary>Fallback HTML document title when PDF metadata does not provide one.</summary>
    [Parameter]
    public string? DocumentTitleFallback { get; set; }

    /// <summary>Optional OfficeIMO PDF to HTML save options.</summary>
    [Parameter]
    public PdfToHtmlOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            PdfToHtmlOptions options = BuildOptions();
            string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
            string html = LoadLogicalDocument(inputPath).ToHtml(options);

            if (!string.IsNullOrWhiteSpace(OutputPath))
            {
                string outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
                if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PDF HTML"))
                {
                    return;
                }

                PdfCommandUtilities.EnsureDirectory(outputPath);
                File.WriteAllText(outputPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                WriteObject(new FileInfo(outputPath));
                return;
            }

            WriteObject(html);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PdfToHtmlFailed", ErrorCategory.InvalidOperation, Path));
        }
    }

    private PdfToHtmlOptions BuildOptions()
    {
        PdfToHtmlOptions options = Options ?? new PdfToHtmlOptions();
        if (Options == null)
        {
            options.Profile = Profile;
            options.ImageExportMode = ImageExportMode;
            options.IncludeMetadata = !NoMetadata.IsPresent;
            options.IncludePageContainers = !NoPageContainers.IsPresent;
            options.IncludeImagePlaceholders = !NoImagePlaceholders.IsPresent;
            options.IncludeLinkAnnotations = IncludeLinkAnnotations.IsPresent;
            options.IncludeFormWidgets = IncludeFormWidgets.IsPresent;
            options.EmitDocumentShell = !Fragment.IsPresent;

            if (MaxEmbeddedImageBytes.HasValue)
            {
                options.MaxEmbeddedImageBytes = MaxEmbeddedImageBytes.Value;
            }

            if (!string.IsNullOrWhiteSpace(DocumentTitleFallback))
            {
                options.DocumentTitleFallback = DocumentTitleFallback!;
            }
        }

        if (!string.IsNullOrWhiteSpace(PageRange))
        {
            options.PageRanges = PdfPageSelection.Parse(PageRange!).Ranges;
        }

        return options;
    }

    private PdfDocumentReadResult LoadLogicalDocument(string path)
    {
        var readOptions = PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent);
        return PdfCommandUtilities.LoadDocument(path, readOptions).Read();
    }
}
