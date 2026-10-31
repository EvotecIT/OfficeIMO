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
public sealed class AddOfficePowerPointListCommand : OfficeCmdlet {
    /// <summary>Gets or sets the list items in display order.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Items")]
    public string[] Items { get; set; } = Array.Empty<string>();
    /// <summary>Gets or sets a block of plain or rich PowerPointListItem commands.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gets or sets whether the list is numbered.</summary>
    [Parameter]
    public SwitchParameter Numbered { get; set; }
    /// <summary>Gets or sets the nesting level, one or two.</summary>
    [Parameter, ValidateRange(1, 2)]
    public int Level { get; set; } = 1;
    /// <summary>Gets or sets whether to emit the native frame backing the first physical list segment.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var scope = PowerPointContentScope.Current;
        if (scope == null) { InvalidArgument(new ArgumentException("Invoke PowerPointList inside PowerPointSlide or PowerPointColumn."), Items); return; }
        var body = scope.CreateTextBody();
        try {
            if (Content != null) {
                using (OfficeDocumentContext.Push(new PowerPointTextScope(scope, list: body.Box))) Content.Invoke();
            } else foreach (var item in Items) body.Box.AddParagraph(item).Level = Level - 1;
            scope.Blocks.Add(new PowerPointLayoutList(Enumerable.Empty<string>(), Numbered) { NativeBody = body });
            if (PassThru) WriteObject(body.Box);
        } catch {
            scope.DiscardTextBody(body);
            throw;
        }
    }
}
