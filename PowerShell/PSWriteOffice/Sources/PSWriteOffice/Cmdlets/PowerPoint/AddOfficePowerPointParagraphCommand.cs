using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a plain or rich paragraph backed by native content before layout.</summary>
/// <example><summary>Capture and edit a designed paragraph.</summary>
/// <code>PowerPointSlide -Title 'Status' {
///     $paragraph = PowerPointParagraph { PowerPointText 'Healthy' -Bold } -PassThru
///     $paragraph.AddText(' in all regions')
/// }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointParagraph", DefaultParameterSetName = "Text")]
[Alias("PowerPointParagraph")]
[OutputType(typeof(PowerPointParagraph))]
public sealed class AddOfficePowerPointParagraphCommand : PowerPointTextPlacementCmdlet {
    /// <summary>Gets or sets the paragraph text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Gets or sets the block of native styled runs.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gets or sets paragraph alignment.</summary>
    [Parameter]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>Gets or sets space before the paragraph; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? SpacingBefore { get; set; }
    /// <summary>Gets or sets space after the paragraph; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? SpacingAfter { get; set; }
    /// <summary>Gets or sets whether plain text is bold.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Bold { get; set; }
    /// <summary>Gets or sets whether plain text is italic.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Italic { get; set; }
    /// <summary>Gets or sets whether plain text has a strike through it.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Strikethrough { get; set; }
    /// <summary>Gets or sets the native underline style.</summary>
    [Parameter(ParameterSetName = "Text")]
    public PowerPointUnderlineStyle? Underline { get; set; }
    /// <summary>Gets or sets the text color or theme role.</summary>
    [Parameter(ParameterSetName = "Text"), OfficeColorTransformation]
    public OfficeColor? Color { get; set; }
    /// <summary>Gets or sets the highlight color.</summary>
    [Parameter(ParameterSetName = "Text"), OfficeColorTransformation]
    public OfficeColor? Highlight { get; set; }
    /// <summary>Gets or sets the font family.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string? Font { get; set; }
    /// <summary>Gets or sets the font size; numbers mean points.</summary>
    [Parameter(ParameterSetName = "Text"), OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }
    /// <summary>Gets or sets whether to emit the live native paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var scope = PowerPointContentScope.Current;
        var textScope = PowerPointTextScope.Current;
        PowerPointTextBox? fixedBox = null;
        PowerPointContentScope? fixedScope = null;
        if (HasFixedPlacement && textScope != null) {
            InvalidArgument(new ArgumentException("Placement is only valid for a direct slide paragraph, not inside text-box, list or table-cell content."), Slide);
            return;
        }
        if (HasFixedPlacement) {
            var placed = CreatePlacedTextBox(string.Empty);
            fixedBox = placed.Box;
            fixedScope = placed.Scope;
            textScope = new PowerPointTextScope(fixedBox.OwnerSlide!.OwnerDocument!, fixedBox);
        }
        if (scope == null && textScope?.HasTextBody != true) { InvalidArgument(new ArgumentException("Supply -Slide or invoke PowerPointParagraph inside a slide, column, text-box or cell content block."), Text); return; }
        var body = textScope?.HasTextBody != true ? scope!.CreateTextBody() : null;
        try {
            var paragraph = body != null ? body.Box.AddParagraph() : textScope!.StartParagraph();
            paragraph.Text = Content == null ? Text : string.Empty;
            if (Alignment.HasValue) paragraph.Alignment = PowerPointTextFormatting.Alignment(Alignment.Value);
            if (SpacingBefore.HasValue) paragraph.SpaceBefore = SpacingBefore.Value;
            if (SpacingAfter.HasValue) paragraph.SpaceAfter = SpacingAfter.Value;
            if (Content != null) {
                paragraph.ClearInlineContent();
                var inline = body != null ? new PowerPointTextScope(scope!, paragraph) : new PowerPointTextScope(textScope!, paragraph);
                using (OfficeDocumentContext.Push(inline.Document))
                using (OfficeDocumentContext.Push(inline)) Content.Invoke();
            } else foreach (var run in paragraph.Runs) {
                if (body != null) PowerPointTextFormatting.Apply(run, scope!, Bold, Italic, Strikethrough, Underline, Color, Highlight, Font, Size);
                else PowerPointTextFormatting.Apply(run, textScope!, Bold, Italic, Strikethrough, Underline, Color, Highlight, Font, Size);
            }
            if (body != null) scope!.Blocks.Add(new PowerPointLayoutText(Text) { NativeBody = body });
            if (fixedBox != null) fixedScope?.TrackFixedFigure(fixedBox);
            if (PassThru) WriteObject(paragraph);
        } catch {
            fixedBox?.Remove();
            if (body != null) scope!.DiscardTextBody(body);
            throw;
        }
    }
}
