using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a native list through the shared layout and continuation rules.</summary>
/// <example><summary>Capture an editable list frame.</summary>
/// <code>$list = PowerPointList 'Availability', 'Capacity' -PassThru
/// $list.Paragraphs[1].Text = 'Capacity planning'</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointList", DefaultParameterSetName = "Items")]
[Alias("PowerPointList")]
[OutputType(typeof(PowerPointTextBox))]
public sealed class AddOfficePowerPointListCommand : PowerPointTextPlacementCmdlet {
    /// <summary>Gets or sets the list items in display order.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Items")]
    public string[] Items { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets a block of plain or rich PowerPointListItem commands.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gets or sets whether the list is numbered.</summary>
    [Parameter]
    public SwitchParameter Numbered { get; set; }
    /// <summary>Gets or sets the top-level numbering scheme; supplying it enables numbering.</summary>
    [Parameter]
    public PowerPointNumberingScheme? Numbering { get; set; }
    /// <summary>Gets or sets a single-character bullet; cannot be combined with numbering.</summary>
    [Parameter, ValidateLength(1, 1)]
    public string? BulletChar { get; set; }
    /// <summary>Gets or sets the nesting level, one or two.</summary>
    [Parameter, ValidateRange(1, 2)]
    public int Level { get; set; } = 1;
    /// <summary>Gets or sets whether to emit the native frame backing the first physical list segment.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var scope = PowerPointContentScope.Current;
        bool fixedPlacement = HasFixedPlacement;
        bool numbered = Numbered || Numbering.HasValue;
        if (BulletChar != null && (numbered || char.IsSurrogate(BulletChar[0]))) {
            InvalidArgument(new ArgumentException("BulletChar must be one non-surrogate character and cannot be combined with numbering."), BulletChar);
            return;
        }
        if (!fixedPlacement && scope == null) { InvalidArgument(new ArgumentException("Supply -Slide or invoke PowerPointList inside PowerPointSlide or PowerPointColumn."), Items); return; }
        PowerPointLayoutNativeText? body = null;
        PowerPointTextBox? box = null;
        PowerPointContentScope? fixedScope = null;
        try {
            if (fixedPlacement) {
                var placed = CreatePlacedTextBox(string.Empty);
                box = placed.Box;
                fixedScope = placed.Scope;
                box.TextBody!.RemoveAllChildren<DocumentFormat.OpenXml.Drawing.Paragraph>();
            } else {
                body = scope!.CreateTextBody();
                box = body.Box;
            }
            int number = 1;
            void ApplyDefaults(PowerPointParagraph paragraph) {
                int level = paragraph.Level ?? 0;
                if (level == 0 && (Numbering.HasValue || fixedPlacement && numbered))
                    paragraph.SetNumbered(Numbering ?? PowerPointNumberingScheme.ArabicPeriod, number++);
                else if (BulletChar != null || fixedPlacement)
                    paragraph.SetBullet(BulletChar?[0] ?? (level == 0 ? '▪' : '–'));
                if (fixedPlacement) PowerPointListParagraphDefaults.Apply(paragraph);
            }
            if (Content != null) {
                var textScope = fixedPlacement
                    ? new PowerPointTextScope(box.OwnerSlide!.OwnerDocument!, box, list: true, listItemDefaults: ApplyDefaults)
                    : new PowerPointTextScope(scope!, list: box, listItemDefaults: ApplyDefaults);
                using (OfficeDocumentContext.Push(textScope.Document))
                using (OfficeDocumentContext.Push(textScope)) Content.Invoke();
            } else foreach (var item in Items) {
                var paragraph = box.AddParagraph(item);
                paragraph.Level = Level - 1;
                ApplyDefaults(paragraph);
            }
            if (fixedPlacement) {
                if (box.Paragraphs.Count == 0) box.AddParagraph(string.Empty);
                fixedScope?.TrackFixedFigure(box);
            } else scope!.Blocks.Add(new PowerPointLayoutList(Enumerable.Empty<string>(), numbered) { NativeBody = body });
            if (PassThru) WriteObject(box);
        } catch {
            if (body != null) scope!.DiscardTextBody(body);
            else box?.Remove();
            throw;
        }
    }
}
