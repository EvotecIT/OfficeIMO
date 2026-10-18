using System;
using System.Management.Automation;
using System.Reflection;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Gets placeholder text boxes from a slide.</summary>
/// <example>
///   <summary>Get the title placeholder.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title</code>
///   <para>Returns the title placeholder textbox if present.</para>
/// </example>
/// <example>
///   <summary>Get placeholders inside a slide DSL block.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx { PptSlide { Get-OfficePowerPointPlaceholder -PlaceholderType Title } }</code>
///   <para>Uses the current slide context.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointPlaceholder")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class GetOfficePowerPointPlaceholderCommand : PSCmdlet
{
    /// <summary>Slide to inspect (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Placeholder type to filter on.</summary>
    [Parameter]
    [Alias("Type")]
    public PowerPointPlaceholderType? PlaceholderType { get; set; }

    /// <summary>Optional placeholder index.</summary>
    [Parameter]
    public uint? Index { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var slide = Slide ?? PowerPointDslContext.Require(this).RequireSlide();

        if (!PlaceholderType.HasValue)
        {
            WriteObject(slide.Placeholders, enumerateCollection: true);
            return;
        }

        var placeholder = slide.GetPlaceholder(PlaceholderType.Value, Index);
        if (placeholder != null)
        {
            WriteObject(placeholder);
        }
    }

}
