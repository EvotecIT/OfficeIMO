using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds an OfficeIMO Word text box to the current Word DSL location.</summary>
/// <example>
///   <summary>Add a positioned report callout.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx {
///     WordTextBox -Text 'Executive summary' -WidthCentimeters 7 -HeightCentimeters 2 -HorizontalOffsetCentimeters 1.5 -VerticalOffsetCentimeters 1 -AutoFitToTextSize
/// }</code>
///   <para>Creates a native OfficeIMO Word text box and applies sizing/positioning through OfficeIMO's text-box API.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTextBox")]
[Alias("WordTextBox")]
[OutputType(typeof(WordTextBox))]
public sealed class AddOfficeWordTextBoxCommand : PSCmdlet
{
    /// <summary>Text to place inside the text box.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Word text wrapping mode.</summary>
    [Parameter]
    public WordImageTextWrapping WrapText { get; set; } = WordImageTextWrapping.Square;

    /// <summary>Width in centimeters.</summary>
    [Parameter]
    [ValidateRange(0.1, double.MaxValue)]
    public double? WidthCentimeters { get; set; }

    /// <summary>Height in centimeters.</summary>
    [Parameter]
    [ValidateRange(0.1, double.MaxValue)]
    public double? HeightCentimeters { get; set; }

    /// <summary>Horizontal offset in centimeters.</summary>
    [Parameter]
    public double? HorizontalOffsetCentimeters { get; set; }

    /// <summary>Vertical offset in centimeters.</summary>
    [Parameter]
    public double? VerticalOffsetCentimeters { get; set; }

    /// <summary>Horizontal alignment for anchored text boxes.</summary>
    [Parameter]
    public WordTextBoxHorizontalAlignment? HorizontalAlignment { get; set; }

    /// <summary>Horizontal relative position anchor.</summary>
    [Parameter]
    public WordHorizontalRelativePosition? HorizontalPositionRelativeFrom { get; set; }

    /// <summary>Vertical relative position anchor.</summary>
    [Parameter]
    public WordVerticalRelativePosition? VerticalPositionRelativeFrom { get; set; }

    /// <summary>Explicit OfficeIMO text-box autofit mode.</summary>
    [Parameter]
    public WordTextBoxAutoFitType? AutoFit { get; set; }

    /// <summary>Resize the text box to fit its text.</summary>
    [Parameter]
    public SwitchParameter AutoFitToTextSize { get; set; }

    /// <summary>Emit the created text box.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = WordDslContext.Require(this);
        var paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
        var textBox = new WordTextBox(context.Document, paragraph, Text, WrapText);

        if (WidthCentimeters.HasValue)
        {
            textBox.WidthCentimeters = WidthCentimeters.Value;
        }

        if (HeightCentimeters.HasValue)
        {
            textBox.HeightCentimeters = HeightCentimeters.Value;
        }

        if (HorizontalOffsetCentimeters.HasValue)
        {
            textBox.HorizontalPositionOffsetCentimeters = HorizontalOffsetCentimeters.Value;
        }

        if (VerticalOffsetCentimeters.HasValue)
        {
            textBox.VerticalPositionOffsetCentimeters = VerticalOffsetCentimeters.Value;
        }

        if (HorizontalAlignment.HasValue)
        {
            textBox.HorizontalAlignment = HorizontalAlignment.Value;
        }

        if (HorizontalPositionRelativeFrom.HasValue)
        {
            textBox.HorizontalPositionRelativeFrom = HorizontalPositionRelativeFrom.Value;
        }

        if (VerticalPositionRelativeFrom.HasValue)
        {
            textBox.VerticalPositionRelativeFrom = VerticalPositionRelativeFrom.Value;
        }

        if (AutoFit.HasValue)
        {
            textBox.AutoFit = AutoFit.Value;
        }

        if (AutoFitToTextSize.IsPresent)
        {
            textBox.AutoFitToTextSize = true;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(textBox);
        }
    }
}
