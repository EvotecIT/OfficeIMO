using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Native ODF paragraph typography for text documents and presentation text boxes.</summary>
public abstract class OfficeOpenDocumentTextCmdlet : OfficeCmdlet {
    /// <summary>Apply bold text; explicit false clears it.</summary>
    [Parameter] public SwitchParameter Bold { get; set; }
    /// <summary>Apply italic text; explicit false clears it.</summary>
    [Parameter] public SwitchParameter Italic { get; set; }
    /// <summary>Apply strikethrough; explicit false clears it.</summary>
    [Parameter] public SwitchParameter Strikethrough { get; set; }
    /// <summary>Native ODF underline line style.</summary>
    [Parameter] public OdfTextDecorationStyle? Underline { get; set; }
    /// <summary>Font family.</summary>
    [Parameter, ValidateNotNullOrEmpty] public string? Font { get; set; }
    /// <summary>Font size; a number means points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength? Size { get; set; }
    /// <summary>Opaque RGB text color, including CSS names and theme roles.</summary>
    [Parameter, OfficeColorTransformation, ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>Paragraph horizontal alignment.</summary>
    [Parameter] public OfficeAlignment? Alignment { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing() {
        if (Size?.ToPoints() <= 0) throw new PSArgumentOutOfRangeException(nameof(Size), Size, "Font size must be positive.");
        if (Color.HasValue && Color.Value.Resolve(OfficeTheme.Default).A != 255)
            throw new PSArgumentException("ODF text color requires an opaque RGB color.", nameof(Color));
    }

    /// <summary>Applies only supplied typography to an ODT paragraph or heading.</summary>
    protected void ApplyTextStyle(OdtParagraph paragraph) {
        if (BoundSwitch(nameof(Bold)) is bool bold) paragraph.Bold = bold;
        if (BoundSwitch(nameof(Italic)) is bool italic) paragraph.Italic = italic;
        if (BoundSwitch(nameof(Strikethrough)) is bool strike) paragraph.Strikethrough = strike;
        if (Underline.HasValue) paragraph.UnderlineStyle = Underline.Value;
        if (Font != null) paragraph.FontFamily = Font;
        if (Size.HasValue) paragraph.FontSize = OdfLength.Points(Size.Value.ToPoints());
        if (Color.HasValue) paragraph.Color = ToNativeColor();
        if (Alignment.HasValue) paragraph.Alignment = Alignment.Value switch {
            OfficeAlignment.Center => OdtParagraphAlignment.Center,
            OfficeAlignment.Right => OdtParagraphAlignment.Right,
            OfficeAlignment.End => OdtParagraphAlignment.End,
            OfficeAlignment.Justify => OdtParagraphAlignment.Justify,
            OfficeAlignment.Start => OdtParagraphAlignment.Start,
            _ => OdtParagraphAlignment.Left
        };
    }

    /// <summary>Applies only supplied typography to an ODP paragraph.</summary>
    protected void ApplyTextStyle(OdpParagraph paragraph) {
        if (BoundSwitch(nameof(Bold)) is bool bold) paragraph.Bold = bold;
        if (BoundSwitch(nameof(Italic)) is bool italic) paragraph.Italic = italic;
        if (BoundSwitch(nameof(Strikethrough)) is bool strike) paragraph.Strikethrough = strike;
        if (Underline.HasValue) paragraph.UnderlineStyle = Underline.Value;
        if (Font != null) paragraph.FontFamily = Font;
        if (Size.HasValue) paragraph.FontSize = OdfLength.Points(Size.Value.ToPoints());
        if (Color.HasValue) paragraph.Color = ToNativeColor();
        if (Alignment.HasValue) paragraph.Alignment = Alignment.Value switch {
            OfficeAlignment.Center => OdpParagraphAlignment.Center,
            OfficeAlignment.Right => OdpParagraphAlignment.Right,
            OfficeAlignment.End => OdpParagraphAlignment.End,
            OfficeAlignment.Justify => OdpParagraphAlignment.Justify,
            OfficeAlignment.Start => OdpParagraphAlignment.Start,
            _ => OdpParagraphAlignment.Left
        };
    }

    private OdfColor ToNativeColor() {
        OfficeColor color = Color!.Value.Resolve(OfficeTheme.Default);
        return new OdfColor(color.R, color.G, color.B);
    }
}
