using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RichEditsUseTheNativeBodyAndItsThemeWithoutChangingNeighbors(bool cell) {
        string path = Path.Combine(_root, "rich-edit.pptx"), inner = Path.Combine(_root, "inner-edit.pptx");
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Dark });
        var slide = document.AddSlide();
        var box = slide.AddTextBox("Old text", OfficeLength.Centimeters(2), OfficeLength.Centimeters(3),
            OfficeLength.Centimeters(8), OfficeLength.Centimeters(2));
        box.Name = "Editable";
        var table = slide.AddTable(2, 2);
        table.GetCell(0, 0).Text = "Keep header";
        table.GetCell(1, 0).Text = "Keep neighbor";
        table.GetCell(1, 1).Text = "Old cell";
        table.GetCell(1, 1).FillColor = "FFFF00";
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("subject", cell ? table : (object)box);
        runspace.SessionStateProxy.SetVariable("inner", inner);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        string edit = cell ? "Set-OfficePowerPointTableCell -Row 2 -Column 2" : "Set-OfficePowerPointShapeText";
        shell.AddScript($$"""
            $changed = $subject | {{edit}} -PassThru -Content {
                PowerPointParagraph {
                    PowerPointText 'Status: ' -Bold -Color Danger
                    New-OfficePowerPoint -Path $inner -Theme Modern {
                        PowerPointSlide { PowerPointParagraph 'Separate document' }
                    }
                    PowerPointText 'Ready' -Italic -Link 'https://example.org/status'
                }
                PowerPointParagraph 'Next review' -SpacingAfter 6
            }
            $changed
            """);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        if (cell) Assert.IsType<PowerPointTableCell>(Assert.Single(result).BaseObject);
        else Assert.Same(box, Assert.Single(result).BaseObject);
        document.Save(path);
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var persistedSlide = Assert.Single(saved.Slides);
        var persistedBox = Assert.Single(persistedSlide.TextBoxes);
        var persistedTable = Assert.Single(persistedSlide.Tables);
        var paragraphs = cell ? persistedTable.GetCell(1, 1).Paragraphs : persistedBox.Paragraphs;
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal("Status: Ready", paragraphs[0].Text);
        Assert.Equal("Next review", paragraphs[1].Text);
        Assert.True(paragraphs[0].Runs[0].Bold);
        Assert.Equal("FF8F8F", paragraphs[0].Runs[0].Color);
        Assert.True(paragraphs[0].Runs[1].Italic);
        Assert.Equal("https://example.org/status", paragraphs[0].Runs[1].Hyperlink!.AbsoluteUri);
        Assert.Equal(6, paragraphs[1].SpaceAfter!.Value.ToPoints());
        Assert.Equal("Keep header", persistedTable.GetCell(0, 0).Text);
        Assert.Equal("Keep neighbor", persistedTable.GetCell(1, 0).Text);
        Assert.Equal("FFFF00", persistedTable.GetCell(1, 1).FillColor);
        Assert.Equal(720000L, persistedBox.Left.ToEmus());
        Assert.Equal(1080000L, persistedBox.Top.ToEmus());
        Assert.Equal(2880000L, persistedBox.Width.ToEmus());
        Assert.Equal(720000L, persistedBox.Height.ToEmus());
        Assert.Equal(cell ? "Old text" : "Old cell", cell ? persistedBox.Text : persistedTable.GetCell(1, 1).Text);
        using var nested = PowerPointDocument.Load(inner);
        Assert.Contains(Assert.Single(nested.Slides).TextBoxes, item => item.Text == "Separate document");
    }

    [Fact]
    public void FixedSpacingUsesTheSelectedSlideBoundsAndCrossAxisAlignmentWhenSaved() {
        string path = Path.Combine(_root, "arranged.pptx");
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var first = slide.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(120), OfficeLength.Points(80), OfficeLength.Points(40), OfficeLength.Points(20));
        var second = slide.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(300), OfficeLength.Points(160), OfficeLength.Points(40), OfficeLength.Points(20));
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("shapes", new object[] { first, second });
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("PowerPointShapeLayout -InputObject $shapes -Slide 1 -Distribute Horizontal -Spacing 1cm -ToSlide -CrossAxisAlignment Top -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(2, result.Count);
        Assert.Same(first, result[0].BaseObject);
        document.Save(path);
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        var shapes = Assert.Single(saved.Slides).Shapes.OrderBy(shape => shape.Left.ToEmus()).ToArray();
        Assert.Equal(2, shapes.Length);
        Assert.Equal(0, shapes[0].Left.ToEmus());
        Assert.Equal(0, shapes[0].Top.ToEmus());
        Assert.Equal(shapes[0].Top, shapes[1].Top);
        Assert.Equal(360000L, shapes[1].Left.ToEmus() - shapes[0].Left.ToEmus() - shapes[0].Width.ToEmus());
        Assert.All(shapes, shape => Assert.Equal(508000L, shape.Width.ToEmus()));
    }
}
