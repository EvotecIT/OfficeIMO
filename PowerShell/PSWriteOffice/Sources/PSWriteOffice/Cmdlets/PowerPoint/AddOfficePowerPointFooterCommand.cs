using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Sets a deck or logical-slide footer slot using text or native inline content.</summary>
/// <para>At the New-/Open- block level the slot applies throughout the deck, including existing slides. Inside PowerPointSlide it overrides that slot for the slide and its continuations. The content block runs once; repeated footers retain native styles, fields and links.</para>
/// <example><summary>Add classification and page numbers.</summary>
/// <code>New-OfficePowerPoint -Path ./Review.pptx {
///     PowerPointFooter 'Contoso | Confidential'
///     PowerPointFooter -Alignment Right { PowerPointSlideNumber -IncludeTotal }
///     PowerPointSlide -Title 'Service review'
/// }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointFooter", DefaultParameterSetName = "Text")]
[Alias("PowerPointFooter")]
public sealed class AddOfficePowerPointFooterCommand : OfficeCmdlet {
    /// <summary>Plain text for the footer slot.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Inline content written with PowerPointText and PowerPointSlideNumber.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }

    /// <summary>Left, Center or Right slot; defaults to Left.</summary>
    [Parameter]
    public OfficeAlignment Alignment { get; set; } = OfficeAlignment.Left;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = OfficeDocumentContext.Find<PowerPointDocument>()
            ?? throw new PSInvalidOperationException("Invoke PowerPointFooter inside New-OfficePowerPoint or Open-OfficePowerPoint.");
        if (PowerPointFooterScope.Current != null) throw new PSInvalidOperationException("A footer cannot contain another footer.");
        if (Alignment is not (OfficeAlignment.Left or OfficeAlignment.Center or OfficeAlignment.Right))
            throw new PSArgumentException("Choose Left, Center or Right for a footer slot.", nameof(Alignment));
        var body = document.CreateFooterTextBody();
        if (Content == null) body.Box.Paragraphs[0].AddRun(Text);
        else {
            using (OfficeDocumentContext.Push(new PowerPointFooterScope(document)))
            using (OfficeDocumentContext.Push(new PowerPointTextScope(document, body.Box))) Content.Invoke();
        }
        var footer = new PowerPointLayoutFooter(body);
        var scope = PowerPointContentScope.Current;
        if (scope != null) scope.Footers[Alignment] = footer;
        else document.SetDeckFooter(Alignment, footer);
    }
}
