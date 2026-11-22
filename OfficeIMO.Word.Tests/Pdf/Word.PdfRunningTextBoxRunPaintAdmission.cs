using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_RunningTextBoxBodyThreeDimensionalPaintRetainsWarnedContent(bool scene) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        Wps.TextBodyProperties body = document.Sections[1].PageSetup.FirstPageHeader.TextBoxes.Single().Anchor!
            .Descendants<Wps.TextBodyProperties>().Single();
        if (scene) body.AddChild(new A.Scene3DType(
            new A.Camera { Preset = A.PresetCameraValues.PerspectiveLeft },
            new A.LightRig { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top }), true);
        else body.AddChild(new A.Shape3DType(new A.BevelTop { Width = 12700L, Height = 12700L }), true);
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document);
    }

}
