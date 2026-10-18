using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a watermark to the current section or header.</summary>
/// <para>Supports text or image watermarks using OfficeIMO.Word.</para>
/// <example>
///   <summary>Add a confidentiality watermark while composing a document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\ProtectedReport.docx {
///     Add-OfficeWordParagraph -Text 'Confidential report'
///     Add-OfficeWordWatermark -Text 'CONFIDENTIAL' -Scale 1.2
///     Protect-OfficeWordDocument -Password 'secret'
/// }</code>
///   <para>Applies a text watermark to the current section and then protects the document through OfficeIMO settings.</para>
/// </example>
/// <example>
///   <summary>Add an image watermark.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\DraftReport.docx {
///     Add-OfficeWordParagraph -Text 'Draft report'
///     Add-OfficeWordWatermark -ImagePath .\Assets\Draft.png -Scale 0.6 -HorizontalOffset 20 -VerticalOffset 40
/// }</code>
///   <para>Uses the image watermark path and placement parameters exposed by OfficeIMO.Word.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordWatermark", DefaultParameterSetName = ParameterSetText)]
[Alias("WordWatermark")]
public sealed class AddOfficeWordWatermarkCommand : PSCmdlet
{
    private const string ParameterSetText = "Text";
    private const string ParameterSetImage = "Image";

    /// <summary>Watermark text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Path to an image watermark.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetImage)]
    public string ImagePath { get; set; } = string.Empty;

    /// <summary>Horizontal offset for the watermark.</summary>
    [Parameter]
    public double? HorizontalOffset { get; set; }

    /// <summary>Vertical offset for the watermark.</summary>
    [Parameter]
    public double? VerticalOffset { get; set; }

    /// <summary>Scale factor for the watermark.</summary>
    [Parameter]
    public double Scale { get; set; } = 1.0;

    /// <summary>Emit the created watermark.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var section = context.RequireSection();
        var header = context.CurrentHeader;

        WordWatermark watermark;
        if (ParameterSetName == ParameterSetImage)
        {
            watermark = header != null
                ? header.AddWatermark(WordWatermarkStyle.Image, ImagePath, HorizontalOffset, VerticalOffset, Scale)
                : section.AddWatermark(WordWatermarkStyle.Image, ImagePath, HorizontalOffset, VerticalOffset, Scale);
        }
        else
        {
            watermark = header != null
                ? header.AddWatermark(WordWatermarkStyle.Text, Text, HorizontalOffset, VerticalOffset, Scale)
                : section.AddWatermark(WordWatermarkStyle.Text, Text, HorizontalOffset, VerticalOffset, Scale);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(watermark);
        }
    }
}
