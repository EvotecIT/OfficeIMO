using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds an OfficeIMO Word text box to the current Word DSL location.</summary>
/// <example>
///   <summary>Add a positioned report callout.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx {
///     WordTextBox -Text 'Executive summary' -Width 7cm -Height 2cm -X 1.5cm -Y 1cm -AutoFitToTextSize
/// }</code>
///   <para>Creates a native OfficeIMO Word text box and applies sizing/positioning through OfficeIMO's text-box API.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTextBox", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordTextBox))]
[Alias("WordTextBox") ]
public sealed class AddOfficeWordTextBoxCommand : PSCmdlet
{
    /// <summary>Text to place inside the text box.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Default")]
    [AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Word text wrapping mode.</summary>
    [Parameter]
    public WordImageTextWrapping Wrap { get; set; } = WordImageTextWrapping.Square;

    /// <summary>Width; numbers mean points and unit suffixes are accepted.</summary>
    [Parameter]
    [PSWriteOffice.Foundation.OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }

    /// <summary>Height; numbers mean points and unit suffixes are accepted.</summary>
    [Parameter]
    [PSWriteOffice.Foundation.OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }

    /// <summary>Horizontal offset; numbers mean points and unit suffixes are accepted.</summary>
    [Parameter]
    [PSWriteOffice.Foundation.OfficeLengthTransformation]
    public OfficeLength? X { get; set; }

    /// <summary>Vertical offset; numbers mean points and unit suffixes are accepted.</summary>
    [Parameter]
    [PSWriteOffice.Foundation.OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }

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
        if (Width.HasValue && Width.Value <= OfficeLength.Zero)
            throw new PSArgumentOutOfRangeException(nameof(Width), Width, "Text-box width must be positive.");
        if (Height.HasValue && Height.Value <= OfficeLength.Zero)
            throw new PSArgumentOutOfRangeException(nameof(Height), Height, "Text-box height must be positive.");
        var context = WordDslContext.Require(this);
        var paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
        var textBox = new WordTextBox(context.Document, paragraph, Text, Wrap);

        if (Width.HasValue)
        {
            textBox.Width = Width.Value;
            textBox.RelativeWidthPercentage = 0;
        }

        if (Height.HasValue)
        {
            textBox.Height = Height.Value;
            textBox.RelativeHeightPercentage = 0;
        }

        if (X.HasValue)
        {
            textBox.HorizontalPositionOffset = X.Value;
        }

        if (Y.HasValue)
        {
            textBox.VerticalPositionOffset = Y.Value;
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
