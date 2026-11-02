using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a native paragraph to the enclosing list, optionally using styled runs.</summary>
/// <example><summary>Write an emphasized list item.</summary>
/// <code>PowerPointList { PowerPointListItem { PowerPointText 'Capacity' -Bold; PowerPointText ': review weekly' } }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointListItem", DefaultParameterSetName = "Text")]
[Alias("PowerPointListItem")]
[OutputType(typeof(PowerPointParagraph))]
public sealed class AddOfficePowerPointListItemCommand : OfficeCmdlet {
    /// <summary>Gets or sets plain item text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Gets or sets the styled-run content block.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gets or sets the one-based list nesting level.</summary>
    [Parameter, ValidateRange(1, 2)]
    public int Level { get; set; } = 1;
    /// <summary>Gets or sets whether plain item text is bold.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Bold { get; set; }
    /// <summary>Gets or sets whether plain item text is italic.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Italic { get; set; }
    /// <summary>Gets or sets whether plain item text has a strike through it.</summary>
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
    /// <summary>Gets or sets whether to emit the live native item paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var scope = PowerPointTextScope.Current;
        if (scope?.List == null) { InvalidArgument(new ArgumentException("Invoke PowerPointListItem inside PowerPointList."), Text); return; }
        var item = scope.List.AddParagraph(Content == null ? Text : string.Empty);
        try {
            item.Level = Level - 1;
            scope.ListItemDefaults?.Invoke(item);
            if (Content != null) {
                item.ClearInlineContent();
                using (OfficeDocumentContext.Push(new PowerPointTextScope(scope, item))) Content.Invoke();
            } else foreach (var run in item.Runs)
                PowerPointTextFormatting.Apply(run, scope, Bold, Italic, Strikethrough, Underline, Color, Highlight, Font, Size);
            if (PassThru) WriteObject(item);
        } catch {
            PowerPointTextRun.DiscardContent(item.Paragraph, item.OwnerPart as DocumentFormat.OpenXml.Packaging.OpenXmlPart);
            throw;
        }
    }
}
