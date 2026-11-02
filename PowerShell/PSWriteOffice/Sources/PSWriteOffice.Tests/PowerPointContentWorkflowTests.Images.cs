using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.ChartForgeX;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    private const string ThreeBandSvg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"300\" height=\"100\" viewBox=\"0 0 300 100\"><rect width=\"100\" height=\"100\" fill=\"red\"/><rect x=\"100\" width=\"100\" height=\"100\" fill=\"green\"/><rect x=\"200\" width=\"100\" height=\"100\" fill=\"blue\"/></svg>";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImagesKeepNativeIdentityPayloadAndFocalCropWithFlowAndTypedPlacement(bool preparedVisual) {
        string source = Path.Combine(_root, "bands[1].svg"), path = Path.Combine(_root, "images.pptx");
        File.WriteAllText(source, ThreeBandSvg);
        var visual = new OfficeVisualSource(ThreeBandSvg) { AlternativeText = "Original visual metadata" }.ToOfficeVisual();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("source", source);
        runspace.SessionStateProxy.SetVariable("target", path);
        runspace.SessionStateProxy.SetVariable("visual", visual);
        string command = preparedVisual ? "$visual | PowerPointImage -AlternativeText 'Three color bands'" : "PowerPointImage -Path $source -AlternativeText 'Three color bands'";
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript($$"""
            New-OfficePowerPoint -Path $target {
                PowerPointSlide -Layout ImageFull -Title 'Evidence' {
                    $script:flowPicture = {{command}} -FocalPoint '0.85,0.5' -PassThru
                }
                $script:flowPicture.Description = 'Edited after layout'
                $script:fixedPicture = {{command}} -Slide 1 -X 2cm -Y 3cm -Width 3cm -Height 3cm -FocalPoint Right -PassThru
                $script:gridPicture = {{command}} -Slide 1 -Column 2 -Span 3 -Y 200pt -Height 3cm -FocalPoint Left -PassThru
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.IsType<PowerPointPicture>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("flowPicture")).BaseObject);
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var slide = Assert.Single(saved.Slides);
        var pictures = slide.Pictures.ToArray();
        Assert.Equal(3, pictures.Length);
        var flow = pictures[0];
        Assert.Equal("Edited after layout", flow.Description);
        Assert.Equal(saved.SlideSize.Width, flow.Width);
        Assert.Equal(saved.SlideSize.Height, flow.Height);
        Assert.InRange(flow.CropLeftRatio, 0.4, 0.42);
        Assert.Equal(0, flow.CropRightRatio, 5);
        var fixedPicture = pictures[1];
        Assert.Equal("Three color bands", fixedPicture.Description);
        Assert.Equal(OfficeLength.Centimeters(2), fixedPicture.Left);
        Assert.Equal(OfficeLength.Centimeters(3), fixedPicture.Top);
        Assert.Equal(OfficeLength.Centimeters(3), fixedPicture.Width);
        Assert.Equal(OfficeLength.Centimeters(3), fixedPicture.Height);
        Assert.Equal(2D / 3D, fixedPicture.CropLeftRatio, 4);
        Assert.Equal(0, fixedPicture.CropRightRatio, 5);
        var grid = pictures[2];
        Assert.True(grid.Left > OfficeTheme.Modern.Spacing.SlideMarginLeft);
        Assert.Equal(OfficeLength.Points(200), grid.Top);
        Assert.Equal(OfficeLength.Centimeters(3), grid.Height);
        Assert.Equal(0, grid.CropLeftRatio, 5);
        Assert.True(grid.CropRightRatio > 0);
        Assert.Equal("Original visual metadata", visual.AlternativeText);
        Assert.All(slide.Pictures, picture => {
            Assert.Equal("image/svg+xml", picture.ContentType);
            Assert.Equal(preparedVisual ? visual.GetPlacementBytes() : File.ReadAllBytes(source), picture.GetImageBytes());
        });
    }

    [Fact]
    public void InvalidImageGridLeavesExistingContentAndRelationshipsIntact() {
        string source = Path.Combine(_root, "bands.svg");
        File.WriteAllText(source, ThreeBandSvg);
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        slide.AddTitle("Retained title");
        int relationships = document.OpenXmlDocument.PresentationPart!.SlideParts.Single().Parts.Count();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("slide", slide);
        runspace.SessionStateProxy.SetVariable("source", source);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("PowerPointImage -Slide $slide -Path $source -Column 10 -Span 4");
        Assert.Empty(shell.Invoke());
        Assert.True(shell.HadErrors);
        Assert.Empty(slide.Pictures);
        Assert.Equal("Retained title", Assert.Single(slide.TextBoxes).Text);
        Assert.Equal(relationships, document.OpenXmlDocument.PresentationPart.SlideParts.Single().Parts.Count());
        Assert.Empty(document.ValidateDocument());
    }
}
