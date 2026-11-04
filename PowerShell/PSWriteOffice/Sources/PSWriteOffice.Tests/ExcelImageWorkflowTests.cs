using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using OfficeIMO;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;
using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelImageWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NaturalTypedAndRangeImagesPersistWithTheirNativeMetadataAndAnchors(bool preparedVisual) {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"300\" height=\"100\"><rect width=\"300\" height=\"100\" fill=\"#147d92\"/></svg>";
        string source = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "[1].svg");
        File.WriteAllText(source, svg);
        try {
            var visual = new OfficeVisualSource(svg) { Title = "Original", AlternativeText = "Original description" }.ToOfficeVisual();
            byte[] payload = visual.GetPlacementBytes();
            using var book = ExcelDocument.Create();
            var sheet = book.AddSheet("Images");
            using var runspace = Open();
            runspace.SessionStateProxy.SetVariable("book", book);
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("visual", visual);
            string input = preparedVisual ? "$visual | ExcelImage -Document $book" : "$book | ExcelImage -Path $source";
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript($$"""
                {{input}} -Sheet 1 -Address B2 -Name Natural -PassThru
                {{input}} -Sheet Images -Address F2 -Width 72 -Height 2cm -OffsetX 3px -AltText 'Availability' -Title 'Placed' -Name Sized -PassThru
                {{input}} -Sheet Images -Address B10:F14 -EndOffsetX 3px -EndOffsetY 5px -Placement MoveOnly -Decorative -Name Range -PassThru
                """);
            var result = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(3, result.Count);
            var natural = Assert.IsType<ExcelImage>(result[0].BaseObject);
            Assert.Equal(OfficeLength.Pixels(300), natural.Width);
            Assert.Equal(OfficeLength.Pixels(100), natural.Height);
            var sized = Assert.IsType<ExcelImage>(result[1].BaseObject);
            Assert.Equal(OfficeLength.Points(72), sized.Width);
            Assert.Equal(OfficeLength.Centimeters(2), sized.Height);
            shell.Commands.Clear();
            shell.AddScript($$"""{{input}} -Sheet Images -Address B10:F14 -Width 72""");
            var error = Record.Exception(() => shell.Invoke());
            Assert.True(error != null || shell.HadErrors);
            Assert.Equal(3, sheet.Images.Count());
            using var saved = new MemoryStream();
            book.Save(saved);
            saved.Position = 0;
            using (var reopened = ExcelDocument.Load(saved)) {
                var images = reopened["Images"].Images.ToArray();
                Assert.Equal(3, images.Length);
                Assert.Equal(OfficeLength.Points(72), images[1].Width);
                Assert.Equal(OfficeLength.Centimeters(2), images[1].Height);
                Assert.Equal("Placed", images[1].Title);
                Assert.Equal("Availability", images[1].Description);
            }
            saved.Position = 0;
            using var package = SpreadsheetDocument.Open(saved, false);
            var drawing = Assert.Single(package.WorkbookPart!.WorksheetParts).DrawingsPart!.WorksheetDrawing!;
            Assert.Equal(2, drawing.Elements<Xdr.OneCellAnchor>().Count());
            var range = Assert.Single(drawing.Elements<Xdr.TwoCellAnchor>());
            Assert.Equal("oneCell", range.EditAs!.InnerText);
            Assert.Equal("28575", range.ToMarker!.ColumnOffset!.Text);
            Assert.Equal("47625", range.ToMarker.RowOffset!.Text);
            Assert.Equal("Original", visual.Title);
            Assert.Equal("Original description", visual.AlternativeText);
            Assert.Equal(payload, visual.GetPlacementBytes());
        } finally {
            File.Delete(source);
        }
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        var type = typeof(AddOfficeExcelImageCommand);
        var command = type.GetCustomAttribute<CmdletAttribute>()!;
        string name = command.VerbName + "-" + command.NounName;
        state.Commands.Add(new SessionStateCmdletEntry(name, type, null));
        state.Commands.Add(new SessionStateAliasEntry("ExcelImage", name));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
