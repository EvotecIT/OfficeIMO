using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PSWriteOffice.Cmdlets.PowerPoint;
using PSWriteOffice.Cmdlets.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.PowerPoint.Html;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed partial class PowerPointOptionWorkflowTests {
    [Theory]
    [InlineData("Pdf", "", true)]
    [InlineData("Pdf", "-NoPictures -NoSpeakerNotes -Font 'Aptos Display'", false)]
    [InlineData("Pdf", "-NoPictures:$false -NoSpeakerNotes:$false", true)]
    [InlineData("Image", "", true)]
    [InlineData("Image", "-NoPictures -NoSlideContent", false)]
    [InlineData("Image", "-NoPictures:$false -NoSlideContent:$false", true)]
    [InlineData("OpenDocument", "", true)]
    [InlineData("OpenDocument", "-NoImages -NoSpeakerNotes -NoBasicFormatting -LossPolicy Block", false)]
    [InlineData("OpenDocument", "-NoImages:$false -NoSpeakerNotes:$false -NoBasicFormatting:$false", true)]
    [InlineData("Html", "", true)]
    [InlineData("Html", "-NoNotes -NoTables -NoHeading -NoSlideHeadings -NoRenderDiagnostics -NoMasterInventory -NoExtractionProof -NoDefaultStyles", false)]
    [InlineData("Html", "-NoNotes:$false -NoTables:$false -NoHeading:$false -NoSlideHeadings:$false -NoRenderDiagnostics:$false -NoMasterInventory:$false -NoExtractionProof:$false -NoDefaultStyles:$false", true)]
    public void NegativeRenderFlagsPreserveDefaultsAndExplicitFalse(string format, string arguments, bool included) {
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("New-OfficePowerPoint" + format + "Options " + arguments);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        object value = Assert.Single(result).BaseObject;
        if (format == "Pdf") {
            var options = Assert.IsType<PowerPointPdfExportOptions>(value);
            Assert.Equal(included, options.IncludePictures);
            Assert.Equal(included, options.IncludeSpeakerNotes);
            if (arguments.Contains("-Font", StringComparison.Ordinal)) Assert.Equal("Aptos Display", options.FontFamily);
        } else if (format == "OpenDocument") {
            var options = Assert.IsType<PowerPointOpenDocumentConversionOptions>(value);
            Assert.Equal(included, options.IncludeImages);
            Assert.Equal(included, options.IncludeSpeakerNotes);
            Assert.Equal(included, options.IncludeBasicFormatting);
            Assert.Equal(arguments.Contains("Block", StringComparison.Ordinal)
                ? OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported : OdfConversionLossPolicy.ReportOnly, options.LossPolicy);
        } else if (format == "Html") {
            var options = Assert.IsType<PowerPointHtmlSaveOptions>(value);
            Assert.Equal(included, options.IncludeNotes);
            Assert.Equal(included, options.IncludeTables);
            Assert.Equal(included, options.IncludeHeading);
            Assert.Equal(included, options.IncludeSlideHeadings);
            Assert.Equal(included, options.IncludeRenderDiagnostics);
            Assert.Equal(included, options.IncludeMasterInventory);
            Assert.Equal(included, options.IncludeExtractionProof);
            Assert.Equal(included, options.IncludeDefaultStyles);
        } else {
            var options = Assert.IsType<PowerPointImageExportOptions>(value);
            Assert.Equal(included, options.IncludePictures);
            Assert.Equal(included, options.IncludeSlideContent);
        }
    }

    [Fact]
    public void ImageResourceBudgetsBindThroughTheirPixelCountNames() {
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("New-OfficePowerPointImageOptions").AddParameter("MaximumRasterPixelCount", 1024L)
            .AddParameter("MaximumTotalRasterPixelCount", 4096L);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var options = Assert.IsType<PowerPointImageExportOptions>(Assert.Single(result).BaseObject);
        Assert.Equal(1024L, options.MaximumRasterPixels);
        Assert.Equal(4096L, options.MaximumTotalRasterPixels);
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficePowerPointPdfOptions", typeof(NewOfficePowerPointPdfOptionsCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficePowerPointImageOptions", typeof(NewOfficePowerPointImageOptionsCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficePowerPointOpenDocumentOptions", typeof(NewOfficePowerPointOpenDocumentOptionsCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficePowerPointHtmlOptions", typeof(NewOfficePowerPointHtmlOptionsCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(PSWriteOffice.Cmdlets.Conversion.ConvertOfficeDocumentCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString()));
}
