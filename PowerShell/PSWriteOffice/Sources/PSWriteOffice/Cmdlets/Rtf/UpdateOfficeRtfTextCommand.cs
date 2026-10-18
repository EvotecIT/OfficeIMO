using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Rtf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Applies lossless text and metadata edits to an RTF document.</summary>
/// <example>
///   <summary>Replace text in an RTF file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeRtf -Path .\Input.rtf -Text 'Status: Draft'
/// Update-OfficeRtfText -Path .\Input.rtf -OutputPath .\Output.rtf -OldText Draft -NewText Final -PassThru</code>
///   <para>Uses OfficeIMO.Rtf's lossless editor to update visible text while preserving untouched RTF syntax.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeRtfText", SupportsShouldProcess = true)]
[Alias("Replace-OfficeRtfText", "RtfText")]
[OutputType(typeof(FileInfo))]
public sealed class UpdateOfficeRtfTextCommand : PSCmdlet
{
    /// <summary>Source RTF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination RTF file path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Visible text to replace.</summary>
    [Parameter]
    public string? OldText { get; set; }

    /// <summary>Replacement visible text.</summary>
    [Parameter]
    public string? NewText { get; set; }

    /// <summary>Use ordinal case-insensitive text replacement.</summary>
    [Parameter]
    public SwitchParameter CaseInsensitive { get; set; }

    /// <summary>Plain paragraphs to append to the end of the RTF document.</summary>
    [Parameter]
    public string[]? AppendParagraph { get; set; }

    /// <summary>Document info fields to set, such as Title, Author, Company, or Comments.</summary>
    [Parameter]
    public IDictionary? DocumentProperty { get; set; }

    /// <summary>Custom user properties to set.</summary>
    [Parameter]
    public IDictionary? UserProperty { get; set; }

    /// <summary>Document variables to set.</summary>
    [Parameter]
    public IDictionary? DocumentVariable { get; set; }

    /// <summary>Emit a FileInfo for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrEmpty(OldText) != string.IsNullOrEmpty(NewText))
        {
            throw new PSArgumentException("Specify both -OldText and -NewText for replacement.");
        }

        var sourcePath = PdfCommandUtilities.ResolvePath(this, Path);
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write updated RTF document"))
        {
            return;
        }

        var editor = RtfDocument.LoadResult(sourcePath).EditLossless();
        if (!string.IsNullOrEmpty(OldText))
        {
            editor.ReplaceText(
                OldText!,
                NewText!,
                CaseInsensitive.IsPresent ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
        }

        if (AppendParagraph != null)
        {
            foreach (var paragraph in AppendParagraph)
            {
                editor.AppendParagraph(paragraph);
            }
        }

        ApplyDocumentProperties(editor);
        ApplyUserProperties(editor);
        ApplyDocumentVariables(editor);

        PdfCommandUtilities.EnsureDirectory(outputPath);
        editor.SaveLossless(outputPath);
        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private void ApplyDocumentProperties(RtfLosslessEditor editor)
    {
        if (DocumentProperty == null)
        {
            return;
        }

        foreach (DictionaryEntry entry in DocumentProperty)
        {
            var fieldName = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(fieldName))
            {
                throw new PSArgumentException("RTF document property names cannot be empty.", nameof(DocumentProperty));
            }

            var field = (RtfDocumentInfoField)Enum.Parse(typeof(RtfDocumentInfoField), fieldName!, ignoreCase: true);
            editor.SetInfo(field, Convert.ToString(entry.Value, CultureInfo.InvariantCulture));
        }
    }

    private void ApplyUserProperties(RtfLosslessEditor editor)
    {
        if (UserProperty == null)
        {
            return;
        }

        foreach (DictionaryEntry entry in UserProperty)
        {
            var name = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new PSArgumentException("RTF user property names cannot be empty.", nameof(UserProperty));
            }

            editor.SetUserProperty(name!, Convert.ToString(entry.Value, CultureInfo.InvariantCulture) ?? string.Empty);
        }
    }

    private void ApplyDocumentVariables(RtfLosslessEditor editor)
    {
        if (DocumentVariable == null)
        {
            return;
        }

        foreach (DictionaryEntry entry in DocumentVariable)
        {
            var name = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new PSArgumentException("RTF document variable names cannot be empty.", nameof(DocumentVariable));
            }

            editor.SetDocumentVariable(name!, Convert.ToString(entry.Value, CultureInfo.InvariantCulture));
        }
    }
}
