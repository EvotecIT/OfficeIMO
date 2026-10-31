using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Visio;
using PSWriteOffice.Cmdlets.Visio;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class VisioLifecycleWorkflowTests {
    [Theory]
    [InlineData("Rectangle", 2D, 1D)]
    [InlineData("Ellipse", 2D, 1D)]
    [InlineData("Diamond", 1D, 1D)]
    [InlineData("TextBox", 2D, 0.5D)]
    public void BasicShapeCommandsBindPhysicalStyleAndRetainTheirNativeDefaultSize(string kind, double width, double height) {
        var document = VisioDocument.Create();
        VisioPage page = document.AddPage("Diagram");
        page.DefaultUnit = VisioMeasurementUnit.Centimeters;
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeVisioRectangle", typeof(AddOfficeVisioRectangleCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeVisioEllipse", typeof(AddOfficeVisioEllipseCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeVisioDiamond", typeof(AddOfficeVisioDiamondCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeVisioTextBox", typeof(AddOfficeVisioTextBoxCommand), null));
        using var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Add-OfficeVisio" + kind).AddParameter("Page", page).AddParameter("Text", "Service")
            .AddParameter("X", "2in").AddParameter("Y", 72).AddParameter("LineWeight", "2pt")
            .AddParameter("Angle", 90).AddParameter("NameU", "service").AddParameter("FillColor", "#123456")
            .AddParameter("LinePattern", 2).AddParameter("FillPattern", 3).AddParameter("PassThru");
        var shape = Assert.IsAssignableFrom<VisioShape>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Same(shape, Assert.Single(page.Shapes));
        Assert.Equal(2, shape.PinX, 6);
        Assert.Equal(1, shape.PinY, 6);
        Assert.Equal(width, shape.Width, 6);
        Assert.Equal(height, shape.Height, 6);
        Assert.Equal(2D / 72D, shape.LineWeight, 6);
        Assert.Equal(Math.PI / 2D, shape.Angle, 6);
        Assert.Equal("service", shape.NameU);
        Assert.Equal(OfficeIMO.OfficeColor.FromRgb(0x12, 0x34, 0x56), shape.FillColor);
        Assert.Equal(2, shape.LinePattern);
        Assert.Equal(3, shape.FillPattern);
    }

    [Fact]
    public void LiveDiagramRetainsPhysicalPageSizeAndDisplayUnitUntilExplicitSave() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-visio-lifecycle-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "Diagram.vsdx");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("New-OfficeVisio", typeof(NewOfficeVisioCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("Save-OfficeVisio", typeof(SaveOfficeVisioCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("New-OfficeVisio").AddParameter("Path", path).AddParameter("Width", "11in")
                .AddParameter("Height", "216mm").AddParameter("DisplayUnit", "Centimeters");
            var document = Assert.IsType<VisioDocument>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.False(File.Exists(path));
            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(11, page.Width, 6);
            Assert.Equal(216 / 25.4, page.Height, 6);
            Assert.Equal(VisioMeasurementUnit.Centimeters, page.DefaultUnit);
            Assert.Equal(VisioMeasurementUnit.Centimeters, page.ScaleMeasurementUnit);
            shell.Commands.Clear();
            shell.AddCommand("Save-OfficeVisio").AddParameter("Document", document).Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            VisioPage saved = Assert.Single(VisioDocument.Load(path).Pages);
            Assert.Equal(11, saved.Width, 6);
            Assert.Equal(216 / 25.4, saved.Height, 6);
            Assert.Equal(VisioMeasurementUnit.Centimeters, saved.ScaleMeasurementUnit);
        } finally {
            Directory.Delete(root, true);
        }
    }
}
