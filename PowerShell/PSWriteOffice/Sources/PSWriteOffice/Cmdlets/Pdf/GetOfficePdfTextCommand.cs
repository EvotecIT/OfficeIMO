using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Extracts text or Markdown from a PDF.</summary>
/// <example>
///   <summary>Extract text and Markdown from a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Get-OfficePdfText -Path .\Examples\Documents\Report.pdf -PageRange '1'
///     Get-OfficePdfText -Path .\Examples\Documents\Report.pdf -AsMarkdown -OutputPath .\Examples\Documents\ReportText.md
/// )
/// $proof</code>
///   <para>Reads plain text directly and writes Markdown readback to a file.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfText", SupportsShouldProcess = true)]
public sealed class GetOfficePdfTextCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional page ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Return logical Markdown instead of plain text.</summary>
    [Parameter]
    public SwitchParameter AsMarkdown { get; set; }

    /// <summary>Return one object per page with PageNumber and Text properties.</summary>
    [Parameter]
    public SwitchParameter ByPage { get; set; }

    /// <summary>Return line-level logical text blocks with page and coordinate metadata.</summary>
    [Parameter]
    public SwitchParameter AsTextBlock { get; set; }

    /// <summary>Password used to extract from a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Optional output text file path.</summary>
    [Parameter]
    public string? OutputPath { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (AsTextBlock.IsPresent && (AsMarkdown.IsPresent || ByPage.IsPresent))
        {
            throw new PSArgumentException("-AsTextBlock cannot be combined with -AsMarkdown or -ByPage.", nameof(AsTextBlock));
        }

        if (ByPage.IsPresent && AsMarkdown.IsPresent)
        {
            throw new PSArgumentException("-ByPage is supported for plain text extraction only.", nameof(ByPage));
        }

        var document = PdfDocument.Load(
            PdfCommandUtilities.ResolvePath(this, Path),
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
        var semanticOptions = string.IsNullOrWhiteSpace(PageRange)
            ? null
            : new PdfReadOptions { PageSelection = PdfPageSelection.Parse(PageRange!) };
        PdfDocumentReadResult result = document.Read(semanticOptions);
        if (AsTextBlock.IsPresent)
        {
            var blocks = result.TextBlocks;

            if (!string.IsNullOrWhiteSpace(OutputPath))
            {
                var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
                if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PDF text blocks"))
                {
                    return;
                }

                PdfCommandUtilities.EnsureDirectory(outputPath);
                File.WriteAllLines(outputPath, blocks.Select(block => block.Text));
                WriteObject(new FileInfo(outputPath));
                return;
            }

            WriteObject(blocks, true);
            return;
        }

        if (ByPage.IsPresent)
        {
            var pages = result.Pages
                .Select(page => new
                {
                    page.PageNumber,
                    Text = string.Join(System.Environment.NewLine, page.TextBlocks.Select(block => block.Text))
                })
                .ToArray();

            if (!string.IsNullOrWhiteSpace(OutputPath))
            {
                var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
                if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PDF page text"))
                {
                    return;
                }

                PdfCommandUtilities.EnsureDirectory(outputPath);
                File.WriteAllText(outputPath, string.Join("\f", pages.Select(page => page.Text)));
                WriteObject(new FileInfo(outputPath));
                return;
            }

            for (var index = 0; index < pages.Length; index++)
            {
                var item = new PSObject();
                item.Properties.Add(new PSNoteProperty("PageNumber", pages[index].PageNumber));
                item.Properties.Add(new PSNoteProperty("Text", pages[index].Text));
                WriteObject(item);
            }

            return;
        }

        var text = AsMarkdown.IsPresent ? result.ToMarkdown() : result.Text;

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
            if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PDF text"))
            {
                return;
            }

            PdfCommandUtilities.EnsureDirectory(outputPath);
            File.WriteAllText(outputPath, text);
            WriteObject(new FileInfo(outputPath));
            return;
        }

        WriteObject(text);
    }
}
