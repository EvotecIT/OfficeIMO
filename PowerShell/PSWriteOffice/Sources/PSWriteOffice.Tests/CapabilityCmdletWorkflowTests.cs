using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Reader;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class CapabilityCmdletWorkflowTests {
    [Fact]
    public void DiscoveredNativeRouteConvertsItsAdvertisedInputAndKeepsCatalogRowsReusable() {
        string root = Directory.CreateTempSubdirectory("officeimo-capability-workflow-").FullName;
        try {
            using var runspace = OpenRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                $route = Get-OfficeDocumentCapability -Conversion -From DOCX -To .PDF -Id DOCX-PDF -Executable
                if (-not $route.ConversionSupported) { throw 'The selected route was not executable' }
                if ($route.ConversionCommand -ne 'Convert-OfficeDocument') { throw 'The route advertised the wrong conversion surface' }
                $route
                Get-OfficeDocumentCapability -Conversion -From Word -To Pdf -Id docx-pdf
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(2, output.Count);
            OfficeConversionCapability route = Assert.IsType<OfficeConversionCapability>(output[0].BaseObject);
            Assert.Equal("docx-pdf", route.Id);
            Assert.Same(route, output[1].BaseObject);
            string source = Path.Combine(root, "source.docx"), target = Path.Combine(root, "source.pdf");
            using (var document = WordDocument.Create()) { document.AddParagraph("Discovered conversion"); document.Save(source); }
            OfficeConverter.Convert(source, target, new OfficeConversionOptions { RouteId = route.Id });
            using var pdf = PdfDocument.Load(target);
            Assert.Contains("Discovered conversion", pdf.Read().Text);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void ProtectionDiscoveryRetainsTypedRowsAndRejectsFilteredCatalogSerialization() {
        using var runspace = OpenRunspace();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $pdf = Get-OfficeDocumentCapability -Protection -Format PDF
            if ($pdf.Id -ne 'pdf-password') { throw 'PDF protection row was not selected' }
            $json = Get-OfficeDocumentCapability -Protection -AsJson | ConvertFrom-Json
            if ($json.id -ne 'OfficeIMO.ProtectedContent') { throw 'The canonical protection catalog was lost' }
            $rejected = $false
            try { Get-OfficeDocumentCapability -Protection -AsJson -Format PDF -ErrorAction Stop }
            catch { $rejected = $_.Exception.Message -like '*cannot be combined with row filters*' }
            if (-not $rejected) { throw 'Filtered whole-catalog output was accepted' }
            $pdf
            """);
        var output = shell.Invoke();
        Assert.Empty(shell.Streams.Error);
        Assert.Equal("pdf-password", Assert.IsType<OfficeIMO.Security.OfficeProtectionCapability>(Assert.Single(output).BaseObject).Id);
    }

    private static Runspace OpenRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(GetOfficeDocumentCapabilityCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
