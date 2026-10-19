using System.Management.Automation;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Services.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Adds a positioned text box to an OpenDocument presentation slide.</summary>
/// <example>
///   <summary>Place a text box using centimetre coordinates.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeOpenDocumentTextBox -Text 'Approved' -X 18 -Y 12 -Width 6 -Height 2</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeOpenDocumentTextBox")]
[OutputType(typeof(OdpTextBox))]
public sealed class AddOfficeOpenDocumentTextBoxCommand : PSCmdlet {
    /// <summary>Slide target. Omit inside Add-OfficeOpenDocumentSlide -Content.</summary>
    [Parameter(ValueFromPipeline = true)]
    public OdpSlide? Slide { get; set; }

    /// <summary>Text box content.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Horizontal position in centimeters.</summary>
    [Parameter] public double X { get; set; } = 1;
    /// <summary>Vertical position in centimeters.</summary>
    [Parameter] public double Y { get; set; } = 1;
    /// <summary>Width in centimeters.</summary>
    [Parameter] [ValidateRange(double.Epsilon, double.MaxValue)] public double Width { get; set; } = 20;
    /// <summary>Height in centimeters.</summary>
    [Parameter] [ValidateRange(double.Epsilon, double.MaxValue)] public double Height { get; set; } = 3;
    /// <summary>Optional shape name.</summary>
    [Parameter] public string? Name { get; set; }
    /// <summary>Emit the created text box.</summary>
    [Parameter] public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OdpSlide slide = Slide ?? OpenDocumentDslContext.Require(this).RequireSlide();
        OdpTextBox textBox = slide.AddTextBox(OdfRect.FromCentimeters(X, Y, Width, Height), Text, Name);
        if (PassThru.IsPresent) WriteObject(textBox);
    }
}
