using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets text in a slide placeholder.</summary>
/// <example>
///   <summary>Set the title placeholder text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\Examples\Documents\PowerPointPlaceholderText.pptx {
///     $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
///     Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Title -Text 'Agenda'
///     Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Body -Text 'Review signals and decisions' -IgnoreMissing
/// }</code>
///   <para>Updates placeholder text when the selected layout exposes matching placeholders.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePowerPointPlaceholderText")]
[OutputType(typeof(PowerPointTextBox))]
[Alias("PptPlaceholderText")]
public sealed class SetOfficePowerPointPlaceholderTextCommand : PSCmdlet
{
    /// <summary>Slide to update (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Placeholder type to target.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Type")]
    public PowerPointPlaceholderType PlaceholderType { get; set; }

    /// <summary>Optional placeholder index.</summary>
    [Parameter]
    public uint? Index { get; set; }

    /// <summary>Text to set.</summary>
    [Parameter(Mandatory = true)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Ignore missing placeholders.</summary>
    [Parameter]
    public SwitchParameter IgnoreMissing { get; set; }

    /// <summary>Emit the placeholder textbox after update.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();
            var placeholder = slide.GetPlaceholder(PlaceholderType, Index);
            if (placeholder == null)
            {
                if (IgnoreMissing.IsPresent)
                {
                    return;
                }

                throw new InvalidOperationException("Placeholder was not found on the slide.");
            }

            placeholder.Text = Text ?? string.Empty;

            if (PassThru.IsPresent)
            {
                WriteObject(placeholder);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointSetPlaceholderTextFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }
}
