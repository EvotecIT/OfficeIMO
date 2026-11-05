using System.Management.Automation;
using OfficeIMO;
using PSWriteOffice.Foundation;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Services.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Adds a positioned text box to an OpenDocument presentation slide.</summary>
/// <example>
///   <summary>Place a text box using centimetre coordinates.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeOpenDocumentTextBox -Text 'Approved' -X 18cm -Y 12cm -Width 6cm -Height 2cm</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeOpenDocumentTextBox", DefaultParameterSetName = "Default")]
[OutputType(typeof(OdpTextBox))]
[Alias("OpenDocumentTextBox") ]
public sealed class AddOfficeOpenDocumentTextBoxCommand : OfficeOpenDocumentTextCmdlet {
    /// <summary>Slide target. Omit inside Add-OfficeOpenDocumentSlide -Content.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Default")]
    public OdpSlide? Slide { get; set; }

    /// <summary>Text box content.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Horizontal position as a physical length; a number means points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength X { get; set; } = OfficeLength.Centimeters(1);
    /// <summary>Vertical position as a physical length; a number means points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength Y { get; set; } = OfficeLength.Centimeters(1);
    /// <summary>Width as a physical length; a number means points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength Width { get; set; } = OfficeLength.Centimeters(20);
    /// <summary>Height as a physical length; a number means points.</summary>
    [Parameter, OfficeLengthTransformation] public OfficeLength Height { get; set; } = OfficeLength.Centimeters(3);
    /// <summary>Optional shape name.</summary>
    [Parameter] public string? Name { get; set; }
    /// <summary>Emit the created text box.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Width.ToPoints() <= 0 || Height.ToPoints() <= 0) throw new PSArgumentOutOfRangeException(nameof(Width), Width, "Text box dimensions must be positive.");
        OdpSlide slide = Slide ?? OpenDocumentDslContext.Require(this).RequireSlide();
        OdpTextBox textBox = slide.AddTextBox(OdfRect.FromCentimeters(X.ToCentimeters(), Y.ToCentimeters(), Width.ToCentimeters(), Height.ToCentimeters()), Text, Name);
        foreach (OdpParagraph paragraph in textBox.Paragraphs) ApplyTextStyle(paragraph);
        if (PassThru.IsPresent) WriteObject(textBox);
    }
}
