using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a live styled run to the enclosing paragraph or list item, or starts a slide paragraph.</summary>
/// <example><summary>Link a status label.</summary>
/// <code>PowerPointParagraph { PowerPointText 'Service status' -Bold -Link 'https://status.contoso.com' }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointText", DefaultParameterSetName = "Text")]
[Alias("PowerPointText")]
[OutputType(typeof(PowerPointTextRun))]
public sealed class AddOfficePowerPointTextCommand : OfficeCmdlet {
    /// <summary>Gets or sets the text to append.</summary>
    [Parameter(Mandatory = true, Position = 0), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Gets or sets bold formatting.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }
    /// <summary>Gets or sets italic formatting.</summary>
    [Parameter]
    public SwitchParameter Italic { get; set; }
    /// <summary>Gets or sets strike-through formatting.</summary>
    [Parameter]
    public SwitchParameter Strikethrough { get; set; }
    /// <summary>Gets or sets the native underline style.</summary>
    [Parameter]
    public PowerPointUnderlineStyle? Underline { get; set; }
    /// <summary>Gets or sets text color or a theme role.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? Color { get; set; }
    /// <summary>Gets or sets the highlight color.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? Highlight { get; set; }
    /// <summary>Gets or sets the font family.</summary>
    [Parameter]
    public string? Font { get; set; }
    /// <summary>Gets or sets the font size; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }
    /// <summary>Gets or sets the hyperlink target.</summary>
    [Parameter]
    public string? Link { get; set; }
    /// <summary>Gets or sets alignment when starting a paragraph directly in the slide or column.</summary>
    [Parameter]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>Gets or sets whether to emit the native run.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var scope = PowerPointContentScope.Current;
        var textScope = PowerPointTextScope.Current;
        if (scope == null && textScope == null) { InvalidArgument(new ArgumentException("Invoke PowerPointText inside a slide or text content block."), Text); return; }
        var paragraph = textScope?.Paragraph;
        if (paragraph == null && textScope?.HasTextBody == true) paragraph = textScope.TextParagraph();
        PowerPointLayoutNativeText? body = null;
        if (paragraph == null) {
            body = scope!.CreateTextBody();
            paragraph = body.Box.AddParagraph();
            paragraph.ClearInlineContent();
            if (Alignment.HasValue) paragraph.Alignment = PowerPointTextFormatting.Alignment(Alignment.Value);
        }
        PowerPointTextRun? run = null;
        try {
            run = paragraph.AddRun(Text);
            if (textScope != null) PowerPointTextFormatting.Apply(run, textScope, Bold, Italic, Strikethrough, Underline, Color, Highlight, Font, Size, Link);
            else PowerPointTextFormatting.Apply(run, scope!, Bold, Italic, Strikethrough, Underline, Color, Highlight, Font, Size, Link);
            if (body != null) scope!.Blocks.Add(new PowerPointLayoutText(Text) { NativeBody = body });
            if (PassThru) WriteObject(run);
        } catch {
            if (body != null) scope!.DiscardTextBody(body);
            else if (run != null) PowerPointTextRun.DiscardContent(run.Run, run.OwnerPart as DocumentFormat.OpenXml.Packaging.OpenXmlPart);
            throw;
        }
    }
}
