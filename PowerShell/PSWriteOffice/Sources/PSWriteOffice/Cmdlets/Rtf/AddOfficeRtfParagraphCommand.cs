using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using OfficeIMO.Rtf;
namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Adds a paragraph to an RTF document.</summary>
/// <example><summary>Use the native OfficeRtfParagraph command.</summary><code>$rtf = New-OfficeRtf
/// Add-OfficeRtfParagraph -Document $rtf -Text 'Maintenance completed.'
/// Save-OfficeRtf -Document $rtf -Path .\Notes\Maintenance.rtf</code></example>
[Cmdlet("Add", "OfficeRtfParagraph", DefaultParameterSetName = "Context")]
[OutputType(typeof(RtfDocument))]
[Alias("RtfParagraph") ]
public sealed class AddOfficeRtfParagraphCommand : OfficeCmdlet {
    /// <summary>Text for this operation.</summary>
    [Parameter(ParameterSetName = "Context", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Document", Mandatory = true, Position = 0)]
    public string? Text { get; set; }
    /// <summary>Bold for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter Bold { get; set; }
    /// <summary>Italic for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter Italic { get; set; }
    /// <summary>Native RTF underline variant.</summary>
    [Parameter] public RtfUnderlineStyle? Underline { get; set; }
    /// <summary>Apply strikethrough; explicit false clears it.</summary>
    [Parameter] public SwitchParameter Strikethrough { get; set; }
    /// <summary>Font family registered in the document font table.</summary>
    [Parameter, ValidateNotNullOrEmpty] public string? Font { get; set; }
    /// <summary>Font size; a number means points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? Size { get; set; }
    /// <summary>Opaque RGB text color, including CSS names and theme roles.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>Hyperlink target written as a native RTF field.</summary>
    [Parameter] public Uri? Link { get; set; }
    /// <summary>Alignment for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>PassThru for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>Document for this operation.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public RtfDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = Document ?? OfficeDocumentContext.Find<RtfDocument>() ?? throw new PSInvalidOperationException("Pass Document or use RtfParagraph inside New-OfficeRtf.");
        if (Size.HasValue && (Size.Value.ToPoints() <= 0 || Size.Value.ToPoints() > int.MaxValue / 2D))
            throw new PSArgumentOutOfRangeException(nameof(Size), Size, "Font size is outside the native RTF range.");
        OfficeColor? color = Color?.Resolve(OfficeTheme.Default);
        if (color.HasValue && color.Value.A != 255) throw new PSArgumentException("RTF text color requires an opaque RGB color.", nameof(Color));
        int? fontId = Font == null ? null : document.AddFont(Font);
        int? colorIndex = color.HasValue ? document.AddColor(color.Value.R, color.Value.G, color.Value.B) : null;
        var paragraph = document.AddParagraph(Text!);
        if (Alignment.HasValue) paragraph.Alignment = Alignment.Value switch {
            OfficeAlignment.Center => RtfTextAlignment.Center,
            OfficeAlignment.Right or OfficeAlignment.End => RtfTextAlignment.Right,
            OfficeAlignment.Justify => RtfTextAlignment.Justify,
            _ => RtfTextAlignment.Left
        };
        foreach (var run in paragraph.Runs) {
            if (BoundSwitch(nameof(Bold)) is bool bold) run.Bold = bold;
            if (BoundSwitch(nameof(Italic)) is bool italic) run.Italic = italic;
            if (BoundSwitch(nameof(Strikethrough)) is bool strike) run.Strikethrough = strike;
            if (Underline.HasValue) run.UnderlineStyle = Underline.Value;
            if (fontId.HasValue) run.FontId = fontId.Value;
            if (Size.HasValue) run.FontSize = Size.Value.ToPoints();
            if (colorIndex.HasValue) run.ForegroundColorIndex = colorIndex.Value;
            if (Link != null) run.Hyperlink = Link;
        }
        if (PassThru) WriteObject(document);
    }
}
