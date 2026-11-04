using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds one styled run inside a PDF paragraph or running header/footer, or a standalone paragraph.</summary>
/// <example><summary>Compose rich PDF text.</summary>
/// <code>PdfParagraph { PdfText 'Review ' -Italic; PdfText 'complete' -Bold -Color Success }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePdfText", DefaultParameterSetName = "Context")]
[OutputType(typeof(PdfDocument))]
[Alias("PdfText") ]
public sealed class AddOfficePdfTextCommand : OfficeCmdlet {
    /// <summary>Live PDF outside a composition block.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public PdfDocument? Document { get; set; }
    /// <summary>Text for this run.</summary>
    [Parameter(Mandatory = true, Position = 0), AllowEmptyString] public string Text { get; set; } = string.Empty;
    /// <summary>Set bold on or off.</summary>
    [Parameter] public SwitchParameter Bold { get; set; }
    /// <summary>Set italic on or off.</summary>
    [Parameter] public SwitchParameter Italic { get; set; }
    /// <summary>Set strikethrough on or off.</summary>
    [Parameter] public SwitchParameter Strikethrough { get; set; }
    /// <summary>Underline pattern; None clears it.</summary>
    [Parameter] public OfficeTextDecorationStyle? Underline { get; set; }
    /// <summary>Run foreground color.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))] public OfficeColor? Color { get; set; }
    /// <summary>Run background color.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))] public OfficeColor? BackgroundColor { get; set; }
    /// <summary>Standard PDF font name or registered embedded family.</summary>
    [Parameter] public string? Font { get; set; }
    /// <summary>Font size; bare numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? Size { get; set; }
    /// <summary>Run baseline.</summary>
    [Parameter] public PdfTextBaseline Baseline { get; set; }
    /// <summary>Leader pattern for an explicit tab text (a single tab character).</summary>
    [Parameter] public PdfTabLeaderStyle TabLeader { get; set; }
    /// <summary>Alignment for an explicit tab text.</summary>
    [Parameter] public PdfTabAlignment TabAlignment { get; set; }
    /// <summary>URI hyperlink target.</summary>
    [Parameter] public string? Link { get; set; }
    /// <summary>Named bookmark target.</summary>
    [Parameter] public string? LinkBookmark { get; set; }
    /// <summary>Link annotation contents.</summary>
    [Parameter] public string? LinkContents { get; set; }
    /// <summary>Standalone paragraph alignment.</summary>
    [Parameter] public OfficeAlignment Alignment { get; set; } = OfficeAlignment.Left;
    /// <summary>Emit the updated live document.</summary>
    [Parameter(ParameterSetName = "Document")] public SwitchParameter PassThru { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        PdfStandardFont? standard = null;
        string? family = Font;
        if (Font != null && System.Enum.TryParse(Font, true, out PdfStandardFont parsed)) { standard = parsed; family = null; }
        var run = new PdfTextRun(Text, color: Color, backgroundColor: BackgroundColor, fontSize: Size,
            standardFont: standard, fontFamily: family, linkUri: Link, linkDestinationName: LinkBookmark,
            linkContents: LinkContents, baseline: Baseline, tabLeader: TabLeader, tabAlignment: TabAlignment);
        if (BoundSwitch(nameof(Bold)) is bool bold) run.SetBold(bold);
        if (BoundSwitch(nameof(Italic)) is bool italic) run.SetItalic(italic);
        if (BoundSwitch(nameof(Strikethrough)) is bool strike) run.SetStrikethrough(strike);
        if (Underline.HasValue) run.SetUnderline(Underline.Value);
        var running = OfficeDocumentContext.Find<PdfRunningTextContent>();
        if (running != null && Document == null) { running.Run(run); return; }
        var runs = OfficeDocumentContext.Find<List<PdfTextRun>>();
        if (runs != null && Document == null) { runs.Add(run); return; }
        var document = PdfCommandUtilities.ComposeContent(this, Document, ParameterSetName, "Document",
            content => content.AddParagraph(paragraph => paragraph.AddRuns(new[] { run }), Alignment, Color));
        if (PassThru && document != null) WriteObject(document);
    }
}
