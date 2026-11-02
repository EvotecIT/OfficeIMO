using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void ShapeQueriesSelectNativeSlidesAndReturnEditableOneBasedPositions() {
        string path = Path.Combine(_root, "shape-query.pptx");
        using var document = PowerPointDocument.Create();
        document.AddSlide().AddTextBox("Keep other slide");
        var slide = document.AddSlide();
        var box = slide.AddTextBox("Needle 42"); box.Name = "Risk.Primary";
        var table = slide.AddTable(1, 1); table.Name = "Risk.Table";
        table.GetCell(0, 0).Text = "Needle in cell";
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        runspace.SessionStateProxy.SetVariable("nativeSlide", slide);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            $literal = @($deck | Get-OfficePowerPointShape -Slide 2 -Text 'needle' -Name 'Risk.*')
            $pattern = @(Get-OfficePowerPointShape -Slide $nativeSlide -Pattern '^Needle in' -Kind Table)
            $position = @(Get-OfficePowerPointShape -Document $deck -Slide 2 -Index 1)
            $exact = @(Get-OfficePowerPointShape -Document $deck -Slide 2 -Text 'needle' -CaseSensitive)
            $literal | Where-Object Kind -eq TextBox | Set-OfficePowerPointShapeText -Text 'Edited needle'
            [pscustomobject]@{ Literal=$literal; Pattern=$pattern; Position=$position; ExactCount=$exact.Count }
            """);
        var row = Assert.Single(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        var literal = Assert.IsType<object[]>(row.Properties["Literal"].Value).Select(value => Assert.IsType<PowerPointShapeInfo>(PSObject.AsPSObject(value).BaseObject)).ToArray();
        Assert.Equal(2, literal.Length);
        Assert.All(literal, info => Assert.Equal(2, info.SlideIndex));
        Assert.Same(box, Assert.Single(literal, info => info.Kind == "TextBox").Shape);
        var pattern = Assert.Single(Assert.IsType<object[]>(row.Properties["Pattern"].Value));
        Assert.Same(table, Assert.IsType<PowerPointShapeInfo>(PSObject.AsPSObject(pattern).BaseObject).Shape);
        var position = Assert.IsType<PowerPointShapeInfo>(PSObject.AsPSObject(Assert.Single(Assert.IsType<object[]>(row.Properties["Position"].Value))).BaseObject);
        Assert.Equal(1, position.ShapeIndex);
        Assert.Same(slide.Shapes[0], position.Shape);
        Assert.Equal(0, row.Properties["ExactCount"].Value);
        document.Save(path);
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal("Keep other slide", Assert.Single(saved.Slides[0].TextBoxes).Text);
        Assert.Equal("Edited needle", Assert.Single(saved.Slides[1].TextBoxes).Text);
        Assert.Equal("Needle in cell", Assert.Single(saved.Slides[1].Tables).GetCell(0, 0).Text);
    }
}
