using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Cmdlets.Pdf;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PdfProofOptionsWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RewriteFactoryControlsRealMetadataPreservationChecks(bool allowMetadataChange) {
        string root = Directory.CreateTempSubdirectory("officeimo-pdf-rewrite-options-").FullName;
        try {
            string source = Path.Combine(root, "original.pdf"), difference = Path.Combine(root, "changed.pdf");
            using (var document = PdfDocument.Create()) {
                document.Compose(builder => builder.Content(content => content.Text("Preserved report content")));
                document.Meta(title: "Original title");
                document.Save(source);
            }
            using (var original = PdfDocument.Load(source)) {
                using var changed = original.UpdateMetadata(CancellationToken.None, title: "Changed title");
                changed.Save(difference);
            }
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("original", source);
            runspace.SessionStateProxy.SetVariable("difference", difference);
            runspace.SessionStateProxy.SetVariable("allowChange", allowMetadataChange);
            var output = Invoke(runspace, """
                $options = New-OfficePdfRewriteOptions -NoMetadata:$allowChange -NoXmpMetadata:$allowChange
                Test-OfficePdfRewrite -ReferencePath $original -DifferencePath $difference -Options $options
                $options
                """);
            var report = Assert.IsType<PdfRewritePreservationReport>(output[0].BaseObject);
            Assert.Equal(allowMetadataChange, report.IsPreserved);
            if (!allowMetadataChange) Assert.Contains(report.Issues, issue => issue.Feature == "Metadata.Title");
            Assert.Equal("Changed title", report.Rewritten.Metadata.Title);
            var settings = Assert.IsType<PdfRewritePreservationOptions>(output[1].BaseObject);
            Assert.Equal(!allowMetadataChange, settings.PreserveMetadata);
            Assert.Null(settings.OriginalReadOptions);
            Assert.Null(settings.RewrittenReadOptions);
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OverlayFactorySelectsSavedSvgLayersWithoutLosingExplicitFalse(bool hideWords) {
        string root = Directory.CreateTempSubdirectory("officeimo-pdf-overlay-options-").FullName;
        try {
            string source = Source(root), destination = Path.Combine(root, "overlay.svg");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            runspace.SessionStateProxy.SetVariable("hideWords", hideWords);
            var output = Invoke(runspace, """
                $options = New-OfficePdfLayoutOverlayOptions -NoWords:$hideWords -NoLines -NoReadingOrder -RegionColor '#FF00FF'
                Export-OfficePdfLayoutOverlay -Path $source -OutputPath $destination -Options $options -PassThru
                """);
            var result = Assert.IsType<OfficeImageExportResult>(Assert.Single(output).BaseObject);
            Assert.Equal(File.ReadAllBytes(destination), result.Bytes);
            var svg = XDocument.Load(destination);
            string[] strokes = svg.Descendants().Attributes("stroke").Select(attribute => attribute.Value.ToUpperInvariant()).ToArray();
            Assert.Contains("#FF00FF", strokes);
            Assert.Equal(!hideWords, strokes.Contains("#2878DC"));
            Assert.DoesNotContain(svg.Descendants(), element => element.Name.LocalName == "text");
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void PngOverlayFactoryAdmitsOnlyOutputsWithinBudget() {
        string root = Directory.CreateTempSubdirectory("officeimo-pdf-overlay-budget-").FullName;
        try {
            string source = Source(root), destination = Path.Combine(root, "absent", "overlay.png");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                $options = New-OfficePdfLayoutOverlayOptions -MaximumRasterPixelCount 1
                try {
                    Export-OfficePdfLayoutOverlay -Path $source -OutputPath $destination -Format Png -Options $options -ErrorAction Stop
                    throw 'Unexpected overlay publication'
                } catch { $_.Exception.Message }
                """);
            var output = shell.Invoke();
            Assert.Contains("pixel limit", Assert.IsType<string>(Assert.Single(output).BaseObject), StringComparison.OrdinalIgnoreCase);
            Assert.False(Directory.Exists(Path.GetDirectoryName(destination)));
            var accepted = Invoke(runspace, """
                $options = New-OfficePdfLayoutOverlayOptions -MaximumRasterPixelCount 40000
                Export-OfficePdfLayoutOverlay -Path $source -OutputPath $destination -Format Png -Scale 0.25 -Options $options -PassThru
                """);
            var image = Assert.IsType<OfficeImageExportResult>(Assert.Single(accepted).BaseObject);
            Assert.Equal(153, image.Width); Assert.Equal(198, image.Height);
            Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, File.ReadAllBytes(destination).Take(8));
            Assert.Equal(File.ReadAllBytes(destination), image.Bytes);
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("Export-OfficePdfLayoutOverlay")]
    [InlineData("Export-OfficePdfXfdf")]
    public void PdfDiagnosticExportsCannotReplaceTheirSourcePdf(string command) {
        string root = Directory.CreateTempSubdirectory("officeimo-pdf-overlay-source-").FullName;
        try {
            string source = Source(root); byte[] before = File.ReadAllBytes(source);
            using var runspace = OpenRunspace(); runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("exportCommand", command);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                try {
                    & $exportCommand -Path $source -OutputPath $source -ErrorAction Stop
                    throw 'Unexpected source replacement'
                } catch { $_.Exception.Message }
                """);
            var output = shell.Invoke();
            Assert.Equal(before, File.ReadAllBytes(source));
            Assert.Contains("different", Assert.IsType<string>(Assert.Single(output).BaseObject), StringComparison.OrdinalIgnoreCase);
        } finally { Directory.Delete(root, true); }
    }

    private static string Source(string root) {
        string path = Path.Combine(root, "source.pdf");
        using var document = PdfDocument.Create();
        document.Compose(builder => builder.Content(content => content.Text("Layout words on one line")));
        document.Save(path);
        return path;
    }

    private static Runspace OpenRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(ExportOfficePdfLayoutOverlayCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open(); return runspace;
    }

    private static PSObject[] Invoke(Runspace runspace, string script) {
        using var shell = PowerShell.Create(); shell.Runspace = runspace; shell.AddScript(script);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        return result.ToArray();
    }
}
