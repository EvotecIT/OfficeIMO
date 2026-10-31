using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Groups typed content into one column of a designed slide.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointColumn")]
[Alias("PowerPointColumn")]
public sealed class AddOfficePowerPointColumnCommand : OfficeCmdlet {
    /// <summary>Gets or sets the column's content block.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public ScriptBlock Content { get; set; } = null!;
    /// <summary>Gets or sets the column title.</summary>
    [Parameter]
    public string? Title { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var parent = PowerPointContentScope.Current;
        if (parent == null) {
            InvalidArgument(new ArgumentException("Invoke PowerPointColumn inside PowerPointSlide."), Content);
            return;
        }
        using var column = new PowerPointContentScope(parent.Document, parent);
        using (OfficeDocumentContext.Push(column)) Content.Invoke();
        try { parent.AcceptColumn(column, Title); }
        catch (ArgumentException error) { InvalidArgument(error, Content); }
    }
}
