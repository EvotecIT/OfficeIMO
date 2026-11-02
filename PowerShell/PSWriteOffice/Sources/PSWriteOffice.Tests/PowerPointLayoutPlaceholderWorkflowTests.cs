using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void LayoutPlaceholderEditsKeepUnboundGeometryAndPersistStyledText() {
        using var document = PowerPointDocument.Create();
        var layout = document.GetSlideLayouts()[0];
        var box = document.EnsureLayoutPlaceholderTextBox(0, layout.LayoutIndex, PowerPointPlaceholderType.Title);
        box.Text = "Template title";
        var original = box.Bounds;
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("layout", layout);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            Set-OfficePowerPointLayoutPlaceholder -Document $deck -TemplateLayout $layout -Type Title `
                -X 1in -Width 20cm -TextMargin 6pt,8pt -Size 31.5pt -Font 'Arial' `
                -Color Accent2 -Highlight '#FFFF00' -Bold -Italic -Underline Double -PassThru
            Set-OfficePowerPointLayoutPlaceholder -Document $deck -TemplateLayout 1 -Type Title -Bold:$false -Italic:$false
            """);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.IsType<PowerPointTextBox>(Assert.Single(result).BaseObject);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        var edited = Assert.IsType<PowerPointTextBox>(saved.GetLayoutPlaceholderTextBox(0, layout.LayoutIndex, PowerPointPlaceholderType.Title));
        Assert.Equal(914400, edited.Left.ToEmus());
        Assert.Equal(7200000, edited.Width.ToEmus());
        Assert.Equal(original.Top, edited.Top);
        Assert.Equal(original.Height, edited.Height);
        Assert.Equal(6, edited.TextMarginTop!.Value.ToPoints());
        Assert.Equal(8, edited.TextMarginRight!.Value.ToPoints());
        var run = Assert.Single(Assert.Single(edited.Paragraphs).Runs);
        Assert.Equal("Template title", run.Text);
        Assert.Equal(31.5, run.FontSize!.Value.ToPoints());
        Assert.Equal("Arial", run.FontName);
        Assert.Equal(saved.GetThemeColor(PowerPointThemeColor.Accent2), run.Color, ignoreCase: true);
        Assert.Equal("FFFF00", run.HighlightColor, ignoreCase: true);
        Assert.False(run.Bold);
        Assert.False(run.Italic);
        Assert.Equal(PowerPointUnderlineStyle.Double, run.UnderlineStyle);
    }

    [Theory]
    [InlineData("X", 60, 40, 600, 120)]
    [InlineData("Y", 30, 60, 600, 120)]
    [InlineData("Width", 30, 40, 60, 120)]
    [InlineData("Height", 30, 40, 600, 60)]
    public void PartialLayoutBoundsPreserveTheInheritedMasterTransform(string parameter, int x, int y, int width, int height) {
        using var document = PowerPointDocument.Create();
        var layout = document.GetSlideLayouts().Single(item => item.Name == "Title and Content");
        var master = document.OpenXmlDocument.PresentationPart!.SlideMasterParts.First();
        var title = master.SlideMaster!.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value
                    == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title);
        title.ShapeProperties!.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D(
            new DocumentFormat.OpenXml.Drawing.Offset { X = 30 * 12700, Y = 40 * 12700 },
            new DocumentFormat.OpenXml.Drawing.Extents { Cx = 600 * 12700, Cy = 120 * 12700 });
        var part = master.SlideLayoutParts.ElementAt(layout.LayoutIndex);
        part.SlideLayout!.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value
                    == DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title).ShapeProperties!.Transform2D = null;
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("layout", layout);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Set-OfficePowerPointLayoutPlaceholder -Document $deck -TemplateLayout $layout -Type Title -" + parameter + " 60pt");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        var edited = saved.GetLayoutPlaceholderTextBox(0, layout.LayoutIndex, PowerPointPlaceholderType.Title)!;
        Assert.Equal(x, edited.Left.ToPoints());
        Assert.Equal(y, edited.Top.ToPoints());
        Assert.Equal(width, edited.Width.ToPoints());
        Assert.Equal(height, edited.Height.ToPoints());
    }

    [Theory]
    [InlineData("4pt", 4, 4, 4, 4)]
    [InlineData("2pt,3pt", 2, 3, 2, 3)]
    [InlineData("1pt,2pt,3pt,4pt", 1, 2, 3, 4)]
    public void LayoutPlaceholderMarginShorthandUsesTopRightBottomLeftOrder(string argument, int top, int right, int bottom, int left) {
        using var document = PowerPointDocument.Create();
        var layout = document.GetSlideLayouts()[0];
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("layout", layout.Name);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Set-OfficePowerPointLayoutPlaceholder -Document $deck -TemplateLayout $layout -Type Title -TextMargin " + argument + " -CreateIfMissing -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var box = Assert.IsType<PowerPointTextBox>(Assert.Single(result).BaseObject);
        Assert.Equal(top, box.TextMarginTop!.Value.ToPoints());
        Assert.Equal(right, box.TextMarginRight!.Value.ToPoints());
        Assert.Equal(bottom, box.TextMarginBottom!.Value.ToPoints());
        Assert.Equal(left, box.TextMarginLeft!.Value.ToPoints());
    }

    [Theory]
    [InlineData("-X 1in -TextMargin 1pt,2pt,3pt")]
    [InlineData("-X 1in -Width -1pt")]
    [InlineData("-X 1in -TextMargin 1pt,2147483648emu")]
    [InlineData("-X 1in -BulletChar '*' -Numbering ArabicPeriod")]
    public void InvalidLayoutPlaceholderArgumentsLeaveExistingNativeBoundsUnchanged(string arguments) {
        using var document = PowerPointDocument.Create();
        var box = document.EnsureLayoutPlaceholderTextBox(0, 0, PowerPointPlaceholderType.Title);
        var original = box.Bounds;
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Set-OfficePowerPointLayoutPlaceholder -Document $deck -TemplateLayout 1 -Type Title " + arguments);
        var failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Equal(ErrorCategory.InvalidArgument, Assert.Single(shell.Streams.Error).CategoryInfo.Category);
        Assert.Equal(original.Left, box.Left);
        Assert.Equal(original.Top, box.Top);
        Assert.Equal(original.Width, box.Width);
        Assert.Equal(original.Height, box.Height);
    }
}
