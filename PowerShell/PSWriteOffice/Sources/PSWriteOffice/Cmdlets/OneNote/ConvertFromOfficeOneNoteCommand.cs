using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.OneNote.Markdown;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.Pdf;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.OneNote;

/// <summary>Converts an offline OneNote section or notebook to semantic Markdown, HTML, or PDF.</summary>
/// <para>Free-form canvas placement and unsupported native data are reported as conversion evidence rather than silently presented as lossless.</para>
/// <example>
///   <summary>Convert a OneNote section to Markdown and inspect fidelity evidence.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$report = ConvertFrom-OfficeOneNote -Path .\Operations.one -OutputPath .\Operations.md -PassThruReport
/// $report | Select-Object HasLoss, Diagnostics</code>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeOneNote", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(OneNoteMarkdownConversionReport), typeof(HtmlConversionReport), typeof(PdfDocumentConversionResult))]
public sealed class ConvertFromOfficeOneNoteCommand : PSCmdlet
{
    /// <summary>Path to a .one section, .onetoc2 notebook index, or .onepkg archive.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination .md, .html, or .pdf path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutPath")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Bounded section and revision-store read options.</summary>
    [Parameter]
    public OneNoteReaderOptions? ReadOptions { get; set; }

    /// <summary>Notebook hierarchy, package, and section-error policy.</summary>
    [Parameter]
    public OneNoteNotebookReaderOptions? NotebookOptions { get; set; }

    /// <summary>OneNote hierarchy, history, and binary-asset projection settings.</summary>
    [Parameter]
    public OneNoteMarkdownOptions? ProjectionOptions { get; set; }

    /// <summary>HTML rendering settings used for .html output.</summary>
    [Parameter]
    public HtmlOptions? HtmlOptions { get; set; }

    /// <summary>Semantic layout and PDF settings used for .pdf output.</summary>
    [Parameter]
    public OneNoteToPdfOptions? PdfOptions { get; set; }

    /// <summary>Fail when the selected projection reports an approximation or omission.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <summary>Overwrite an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Return conversion evidence instead of file information.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        var extension = System.IO.Path.GetExtension(output).ToLowerInvariant();
        if (extension != ".md" && extension != ".html" && extension != ".pdf")
        {
            throw new PSArgumentException("OutputPath must use the .md, .html, or .pdf extension.", nameof(OutputPath));
        }

        if (HtmlOptions != null && extension != ".html")
        {
            throw new PSArgumentException("HtmlOptions can only be used with .html output.", nameof(HtmlOptions));
        }

        if (PdfOptions != null && extension != ".pdf")
        {
            throw new PSArgumentException("PdfOptions can only be used with .pdf output.", nameof(PdfOptions));
        }

        if (extension == ".pdf" && PdfOptions != null && ProjectionOptions != null)
        {
            throw new PSArgumentException("Specify OneNote projection settings either through -ProjectionOptions or through -PdfOptions, not both.");
        }

        if (File.Exists(output) && !Force.IsPresent)
        {
            throw new IOException($"File '{output}' already exists. Use -Force to overwrite it.");
        }

        if (!ShouldProcess(output, $"Convert OneNote to {extension.TrimStart('.').ToUpperInvariant()}"))
        {
            return;
        }

        var source = OneNoteCommandUtilities.Read(input, ReadOptions, NotebookOptions);
        var evidence = extension switch
        {
            ".md" => WriteMarkdown(source, output),
            ".html" => WriteHtml(source, output),
            ".pdf" => WritePdf(source, output),
            _ => throw new InvalidOperationException("Unsupported OneNote output format.")
        };
        WriteObject(PassThruReport.IsPresent ? evidence : new FileInfo(output));
    }

    private object WriteMarkdown(object source, string output)
    {
        var result = source switch
        {
            OneNoteSection section => section.ToMarkdownDocumentResult(ProjectionOptions),
            OneNoteNotebook notebook => notebook.ToMarkdownDocumentResult(ProjectionOptions),
            _ => throw new InvalidOperationException("Unsupported OneNote source model.")
        };
        if (FailOnLoss.IsPresent) result.RequireNoLoss();
        AtomicFileWriter.Write(output, new UTF8Encoding(false).GetBytes(result.Value.ToMarkdown()), Force.IsPresent);
        return result.Report;
    }

    private object WriteHtml(object source, string output)
    {
        var result = source switch
        {
            OneNoteSection section => section.ToHtmlDocumentResult(ProjectionOptions, HtmlOptions),
            OneNoteNotebook notebook => notebook.ToHtmlDocumentResult(ProjectionOptions, HtmlOptions),
            _ => throw new InvalidOperationException("Unsupported OneNote source model.")
        };
        if (FailOnLoss.IsPresent) result.RequireNoLoss();
        AtomicFileWriter.Write(output, new UTF8Encoding(false).GetBytes(result.Value), Force.IsPresent);
        return result.Report;
    }

    private object WritePdf(object source, string output)
    {
        var options = PdfOptions ?? new OneNoteToPdfOptions { ProjectionOptions = ProjectionOptions ?? new OneNoteMarkdownOptions() };
        var result = source switch
        {
            OneNoteSection section => section.ToPdfDocumentResult(options),
            OneNoteNotebook notebook => notebook.ToPdfDocumentResult(options),
            _ => throw new InvalidOperationException("Unsupported OneNote source model.")
        };
        AtomicFileWriter.Write(output, Force.IsPresent, temporaryPath =>
        {
            result.Save(temporaryPath).RequireSuccess();
            if (FailOnLoss.IsPresent) result.RequireNoLoss();
        });
        return result;
    }
}
