using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets or clears color, image and shape backgrounds in the enclosing PDF composition.</summary>
[Cmdlet(VerbsCommon.Set, "OfficePdfBackground", DefaultParameterSetName = "Color")]
[Alias("PdfBackground")]
public sealed class SetOfficePdfBackgroundCommand : PSCmdlet {
    /// <summary>Background color.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Color"), OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor Color { get; set; }
    /// <summary>Literal background image path.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Image")]
    public string ImagePath { get; set; } = string.Empty;
    /// <summary>How the background image fits the page.</summary>
    [Parameter(ParameterSetName = "Image")]
    public OfficeImageFit Fit { get; set; } = OfficeImageFit.Cover;
    /// <summary>Background image opacity.</summary>
    [Parameter(ParameterSetName = "Image"), ValidateRange(0D, 1D)]
    public double? Opacity { get; set; }
    /// <summary>Clear the selected background decoration.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Clear")]
    public SwitchParameter Clear { get; set; }
    /// <summary>Background decorations to clear.</summary>
    [Parameter(ParameterSetName = "Clear")]
    public OfficePdfBackgroundTarget Target { get; set; } = OfficePdfBackgroundTarget.All;
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = PdfDslContext.Require(this);
        if (ParameterSetName == "Image") {
            var bytes = PdfImageDocumentSource.FromFile(PdfCommandUtilities.ResolvePath(this, ImagePath)).GetBytes();
            context.ConfigurePage(page => page.BackgroundImage(bytes, Fit, Opacity));
        } else if (ParameterSetName == "Clear") {
            context.ConfigurePage(page => {
                if (Target == OfficePdfBackgroundTarget.All || Target == OfficePdfBackgroundTarget.Color) page.Background(null);
                if (Target == OfficePdfBackgroundTarget.All || Target == OfficePdfBackgroundTarget.Image) page.BackgroundImage(null);
                if (Target == OfficePdfBackgroundTarget.All || Target == OfficePdfBackgroundTarget.Shapes) page.ClearBackgroundShapes();
            });
        } else context.ConfigurePage(page => page.Background(Color));
    }
}
