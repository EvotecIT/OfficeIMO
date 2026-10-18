using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Rtf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Converts Word, HTML, PDF, or Markdown input to RTF.</summary>
/// <example>
///   <summary>Convert Word to RTF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx { WordParagraph -Text 'Summary' }
/// ConvertTo-OfficeRtf -WordPath .\Report.docx -OutputPath .\Report.rtf -PassThru</code>
///   <para>Loads the Word document and saves an RTF file using OfficeIMO.Word.Rtf.</para>
/// </example>
/// <example>
///   <summary>Convert HTML to RTF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeRtf -Html '&lt;h1&gt;Report&lt;/h1&gt;' -OutputPath .\Report.rtf</code>
///   <para>Creates a Word document from HTML and serializes it to RTF.</para>
/// </example>
/// <example>
///   <summary>Convert PDF to RTF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeRtf -PdfPath .\Report.pdf -OutputPath .\Report.rtf</code>
///   <para>Uses OfficeIMO.Rtf.Pdf's semantic PDF reader to write RTF output.</para>
/// </example>
/// <example>
///   <summary>Convert Markdown to RTF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeRtf -MarkdownPath .\Report.md -OutputPath .\Report.rtf -PassThru</code>
///   <para>Parses Markdown and writes RTF using OfficeIMO.Rtf.Markdown.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeRtf", DefaultParameterSetName = ParameterSetWordPath, SupportsShouldProcess = true)]
[Alias("ConvertTo-Rtf")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeRtfCommand : PSCmdlet
{
    private const string ParameterSetWordPath = "WordPath";
    private const string ParameterSetWordDocument = "WordDocument";
    private const string ParameterSetHtml = "Html";
    private const string ParameterSetHtmlPath = "HtmlPath";
    private const string ParameterSetPdfPath = "PdfPath";
    private const string ParameterSetMarkdown = "Markdown";
    private const string ParameterSetMarkdownPath = "MarkdownPath";

    /// <summary>Path to a .docx file to convert to RTF.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetWordPath)]
    public string WordPath { get; set; } = string.Empty;

    /// <summary>Word document instance to convert to RTF.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetWordDocument)]
    public WordDocument WordDocument { get; set; } = null!;

    /// <summary>HTML markup to convert to RTF through the Word HTML converter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetHtml)]
    public string Html { get; set; } = string.Empty;

    /// <summary>Path to an HTML file to convert to RTF.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetHtmlPath)]
    public string HtmlPath { get; set; } = string.Empty;

    /// <summary>Path to a PDF file to convert to semantic RTF.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPdfPath)]
    public string PdfPath { get; set; } = string.Empty;

    /// <summary>Markdown text to convert to RTF.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetMarkdown)]
    public string Markdown { get; set; } = string.Empty;

    /// <summary>Path to a Markdown file to convert to RTF.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetMarkdownPath)]
    public string MarkdownPath { get; set; } = string.Empty;

    /// <summary>Optional destination RTF path. When omitted, raw RTF text is returned.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Optional font family for HTML to Word conversion before RTF serialization.</summary>
    [Parameter(ParameterSetName = ParameterSetHtml)]
    [Parameter(ParameterSetName = ParameterSetHtmlPath)]
    public string? FontFamily { get; set; }

    /// <summary>Base path used to resolve relative HTML resources.</summary>
    [Parameter(ParameterSetName = ParameterSetHtml)]
    [Parameter(ParameterSetName = ParameterSetHtmlPath)]
    public string? BasePath { get; set; }

    /// <summary>Paths to CSS stylesheets to apply during HTML conversion.</summary>
    [Parameter(ParameterSetName = ParameterSetHtml)]
    [Parameter(ParameterSetName = ParameterSetHtmlPath)]
    public string[]? StylesheetPath { get; set; }

    /// <summary>Inline CSS stylesheets to apply during HTML conversion.</summary>
    [Parameter(ParameterSetName = ParameterSetHtml)]
    [Parameter(ParameterSetName = ParameterSetHtmlPath)]
    public string[]? StylesheetContent { get; set; }

    /// <summary>Preserve raw HTML Markdown blocks as plain text in the RTF output.</summary>
    [Parameter(ParameterSetName = ParameterSetMarkdown)]
    [Parameter(ParameterSetName = ParameterSetMarkdownPath)]
    public SwitchParameter PreserveRawHtmlAsText { get; set; }

    /// <summary>Emit a FileInfo when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            if (ParameterSetName == ParameterSetPdfPath)
            {
                ConvertPdf();
                return;
            }

            if (ParameterSetName == ParameterSetMarkdown || ParameterSetName == ParameterSetMarkdownPath)
            {
                ConvertMarkdown();
                return;
            }

            WordDocument? document = null;
            var dispose = false;
            try
            {
                document = LoadWordDocument(out dispose);
                WriteRtf(document);
            }
            finally
            {
                if (dispose)
                {
                    document?.Dispose();
                }
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ConvertToOfficeRtfFailed", ErrorCategory.InvalidOperation, GetErrorTarget()));
        }
    }

    private WordDocument LoadWordDocument(out bool dispose)
    {
        dispose = false;
        if (ParameterSetName == ParameterSetWordDocument)
        {
            return WordDocument;
        }

        dispose = true;
        if (ParameterSetName == ParameterSetWordPath)
        {
            return WordDocumentService.LoadDocument(PdfCommandUtilities.ResolvePath(this, WordPath), readOnly: true, autoSave: false);
        }

        var html = Html;
        string? htmlDirectory = null;
        if (ParameterSetName == ParameterSetHtmlPath)
        {
            var resolvedHtml = PdfCommandUtilities.ResolvePath(this, HtmlPath);
            html = File.ReadAllText(resolvedHtml);
            htmlDirectory = Path.GetDirectoryName(resolvedHtml);
        }

        if (string.IsNullOrWhiteSpace(html))
        {
            throw new PSArgumentException("HTML content cannot be empty.", nameof(Html));
        }

        return HtmlConversionDocument.Parse(html).ToWordDocument(CreateHtmlToWordOptions(htmlDirectory));
    }

    private HtmlToWordOptions CreateHtmlToWordOptions(string? htmlDirectory)
    {
        var options = new HtmlToWordOptions();
        if (!string.IsNullOrWhiteSpace(FontFamily))
        {
            options.FontFamily = FontFamily;
        }

        if (!string.IsNullOrWhiteSpace(BasePath))
        {
            options.BasePath = PdfCommandUtilities.ResolvePath(this, BasePath!);
        }
        else if (!string.IsNullOrWhiteSpace(htmlDirectory))
        {
            options.BasePath = htmlDirectory;
        }

        if (StylesheetPath != null)
        {
            foreach (var entry in StylesheetPath)
            {
                if (!string.IsNullOrWhiteSpace(entry))
                {
                    options.StylesheetPaths.Add(PdfCommandUtilities.ResolvePath(this, entry));
                }
            }
        }

        if (StylesheetContent != null)
        {
            foreach (var entry in StylesheetContent)
            {
                if (!string.IsNullOrWhiteSpace(entry))
                {
                    options.StylesheetContents.Add(entry);
                }
            }
        }

        return options;
    }

    private void WriteRtf(WordDocument document)
    {
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(document.ToRtf());
            return;
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write RTF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        document.SaveAsRtf(outputPath, encoding: new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private void ConvertPdf()
    {
        var sourcePath = PdfCommandUtilities.ResolvePath(this, PdfPath);
        var document = PdfCommandUtilities.LoadDocument(sourcePath).Read().ToRtfDocument();
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(document.ToRtf());
            return;
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write RTF converted from PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        document.Save(outputPath, encoding: new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private void ConvertMarkdown()
    {
        var markdown = Markdown;
        if (ParameterSetName == ParameterSetMarkdownPath)
        {
            markdown = File.ReadAllText(PdfCommandUtilities.ResolvePath(this, MarkdownPath));
        }

        if (string.IsNullOrWhiteSpace(markdown))
        {
            throw new PSArgumentException("Markdown content cannot be empty.", nameof(Markdown));
        }

        var options = new MarkdownToRtfOptions
        {
            PreserveRawHtmlAsText = PreserveRawHtmlAsText.IsPresent
        };

        var rtf = MarkdownReader.Parse(markdown, options.ReaderOptions).ToRtfDocument(options).ToRtf();
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(rtf);
            return;
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write RTF converted from Markdown"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        File.WriteAllText(outputPath, rtf, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private object GetErrorTarget()
    {
        return ParameterSetName switch
        {
            ParameterSetWordPath => WordPath,
            ParameterSetWordDocument => WordDocument,
            ParameterSetHtml => Html,
            ParameterSetHtmlPath => HtmlPath,
            ParameterSetPdfPath => PdfPath,
            ParameterSetMarkdown => Markdown,
            ParameterSetMarkdownPath => MarkdownPath,
            _ => this
        };
    }
}
