using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Rtf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Converts RTF input to Word, HTML, PDF, or Markdown output.</summary>
/// <example>
///   <summary>Convert RTF to Word.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeRtf -Path .\Report.rtf -Text 'Summary', 'Ready for review'
/// ConvertFrom-OfficeRtf -Path .\Report.rtf -As Word -OutputPath .\Report.docx -PassThru</code>
///   <para>Loads the RTF file and saves a Word document using OfficeIMO.Word.Rtf.</para>
/// </example>
/// <example>
///   <summary>Convert RTF to HTML.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertFrom-OfficeRtf -Path .\Report.rtf -As Html -OutputPath .\Report.html</code>
///   <para>Converts RTF to Word, then serializes Word to HTML.</para>
/// </example>
/// <example>
///   <summary>Convert RTF to PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertFrom-OfficeRtf -Path .\Report.rtf -As Pdf -OutputPath .\Report.pdf</code>
///   <para>Uses OfficeIMO.Rtf.Pdf to save a PDF file.</para>
/// </example>
/// <example>
///   <summary>Convert RTF to Markdown.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertFrom-OfficeRtf -Path .\Report.rtf -As Markdown -OutputPath .\Report.md -PassThru</code>
///   <para>Converts the RTF document to Markdown using OfficeIMO.Rtf.Markdown.</para>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeRtf", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("ConvertFrom-Rtf")]
[OutputType(typeof(WordDocument), typeof(string), typeof(PdfDocument), typeof(FileInfo))]
public sealed class ConvertFromOfficeRtfCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>RTF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Raw RTF text.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Target document format.</summary>
    [Parameter(Mandatory = true)]
    public OfficeRtfConversionTarget As { get; set; }

    /// <summary>Optional output path. When omitted, the converted object or text is returned.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Optional font family for RTF to HTML conversion.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Include font styles as inline CSS for RTF to HTML conversion.</summary>
    [Parameter]
    public SwitchParameter IncludeFontStyles { get; set; }

    /// <summary>Include list style metadata for RTF to HTML conversion.</summary>
    [Parameter]
    public SwitchParameter IncludeListStyles { get; set; }

    /// <summary>Emit paragraph styles as CSS classes for RTF to HTML conversion.</summary>
    [Parameter]
    public SwitchParameter IncludeParagraphClasses { get; set; }

    /// <summary>Emit run styles as CSS classes for RTF to HTML conversion.</summary>
    [Parameter]
    public SwitchParameter IncludeRunClasses { get; set; }

    /// <summary>Include the built-in default CSS for RTF to HTML conversion.</summary>
    [Parameter]
    public SwitchParameter IncludeDefaultCss { get; set; }

    /// <summary>Store image references as file paths instead of base64 data URIs for HTML output.</summary>
    [Parameter]
    public SwitchParameter UseImagePaths { get; set; }

    /// <summary>Include hidden RTF text when converting to PDF or Markdown.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenText { get; set; }

    /// <summary>Exclude RTF images from PDF output.</summary>
    [Parameter]
    public SwitchParameter ExcludeImages { get; set; }

    /// <summary>Exclude RTF tables from PDF output.</summary>
    [Parameter]
    public SwitchParameter ExcludeTables { get; set; }

    /// <summary>Exclude RTF headers and footers from PDF output.</summary>
    [Parameter]
    public SwitchParameter ExcludeHeaderFooters { get; set; }

    /// <summary>Exclude RTF notes from PDF output.</summary>
    [Parameter]
    public SwitchParameter ExcludeNotes { get; set; }

    /// <summary>Do not emit HTML comments for unsupported RTF features when converting to Markdown.</summary>
    [Parameter]
    public SwitchParameter NoUnsupportedHtmlComments { get; set; }

    /// <summary>Emit a FileInfo when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            switch (As)
            {
                case OfficeRtfConversionTarget.Word:
                    ConvertToWord();
                    break;
                case OfficeRtfConversionTarget.Html:
                    ConvertToHtml();
                    break;
                case OfficeRtfConversionTarget.Pdf:
                    ConvertToPdf();
                    break;
                case OfficeRtfConversionTarget.Markdown:
                    ConvertToMarkdown();
                    break;
                default:
                    throw new PSArgumentOutOfRangeException(nameof(As), As, "Unsupported RTF conversion target.");
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ConvertFromOfficeRtfFailed", ErrorCategory.InvalidOperation,
                ParameterSetName == ParameterSetPath ? Path : Text));
        }
    }

    private void ConvertToWord()
    {
        var document = LoadWordDocument();
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(document);
            return;
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write Word document converted from RTF"))
        {
            document.Dispose();
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        try
        {
            document.Save(outputPath);
        }
        finally
        {
            document.Dispose();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private void ConvertToHtml()
    {
        using var document = LoadWordDocument();
        var options = CreateWordToHtmlOptions();
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(document.ToHtml(options));
            return;
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write HTML converted from RTF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        document.SaveAsHtml(outputPath, options);
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private void ConvertToPdf()
    {
        var options = new RtfToPdfOptions
        {
            IncludeHiddenText = IncludeHiddenText.IsPresent,
            IncludeImages = !ExcludeImages.IsPresent,
            IncludeTables = !ExcludeTables.IsPresent,
            IncludeHeaderFooters = !ExcludeHeaderFooters.IsPresent,
            IncludeNotes = !ExcludeNotes.IsPresent
        };

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
            if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PDF converted from RTF"))
            {
                return;
            }

            PdfCommandUtilities.EnsureDirectory(outputPath);
            LoadRtfDocument().SaveAsPdf(outputPath, options).RequireSuccess();

            if (PassThru.IsPresent)
            {
                WriteObject(new FileInfo(outputPath));
            }

            return;
        }

        var document = LoadRtfDocument().ToPdfDocument(options);
        WriteObject(document);
    }

    private void ConvertToMarkdown()
    {
        var options = new RtfToMarkdownOptions
        {
            IncludeHiddenText = IncludeHiddenText.IsPresent,
            EmitUnsupportedHtmlComments = !NoUnsupportedHtmlComments.IsPresent
        };

        var markdown = LoadRtfDocument().ToMarkdown(options);
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteObject(markdown);
            return;
        }

        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write Markdown converted from RTF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        File.WriteAllText(outputPath, markdown, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private RtfDocument LoadRtfDocument()
    {
        return ParameterSetName == ParameterSetPath
            ? RtfDocument.Load(PdfCommandUtilities.ResolvePath(this, Path), encoding: new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))
            : RtfDocument.Read(Text).Document;
    }

    private WordDocument LoadWordDocument()
    {
        return LoadRtfDocument().ToWordDocument();
    }

    private WordToHtmlOptions CreateWordToHtmlOptions()
    {
        var options = new WordToHtmlOptions
        {
            IncludeFontStyles = IncludeFontStyles.IsPresent,
            IncludeListStyles = IncludeListStyles.IsPresent,
            IncludeParagraphClasses = IncludeParagraphClasses.IsPresent,
            IncludeRunClasses = IncludeRunClasses.IsPresent,
            IncludeDefaultCss = IncludeDefaultCss.IsPresent
        };

        if (!string.IsNullOrWhiteSpace(FontFamily))
        {
            options.FontFamily = FontFamily;
        }

        if (UseImagePaths.IsPresent)
        {
            options.EmbedImagesAsBase64 = false;
        }

        return options;
    }
}
