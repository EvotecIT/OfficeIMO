using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Converts PDF logical text readback to Markdown.</summary>
/// <example>
///   <summary>Export logical PDF text as Markdown.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$markdownPath = '.\Examples\Documents\Report.md'
/// ConvertTo-OfficePdfMarkdown -Path .\Examples\Documents\Report.pdf -PageRange '1-3' -OutputPath $markdownPath
/// Get-Content $markdownPath -TotalCount 20</code>
///   <para>Writes Markdown readback for selected pages to a file.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfMarkdown", SupportsShouldProcess = true)]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficePdfMarkdownCommand : PSCmdlet
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

    /// <summary>Optional output Markdown file path.</summary>
    [Parameter]
    public string? OutputPath { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var readOptions = string.IsNullOrWhiteSpace(PageRange)
            ? null
            : new PdfReadOptions { PageSelection = PdfPageSelection.Parse(PageRange!) };
        var document = PdfDocument.Load(
            PdfCommandUtilities.ResolvePath(this, Path),
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
        var markdown = document.Read(readOptions).ToMarkdown();

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
            if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PDF Markdown"))
            {
                return;
            }

            PdfCommandUtilities.EnsureDirectory(outputPath);
            File.WriteAllText(outputPath, markdown);
            WriteObject(new FileInfo(outputPath));
            return;
        }

        WriteObject(markdown);
    }
}
