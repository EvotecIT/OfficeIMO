using System.Management.Automation;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void PdfHtmlFactoryFeedsNativeShellAndMetadataOptions(bool fragment, bool hideMetadata) {
        using var document = PdfDocument.Create();
        document.Compose(builder => builder.Content(content => content.Text("Quarterly report content")));
        document.Meta(author: "Report author");
        using var runspace = OpenRunspace();
        runspace.SessionStateProxy.SetVariable("document", document);
        runspace.SessionStateProxy.SetVariable("fragment", fragment);
        runspace.SessionStateProxy.SetVariable("hideMetadata", hideMetadata);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            $options = New-OfficePdfHtmlOptions -Fragment:$fragment -NoMetadata:$hideMetadata -DocumentTitleFallback 'Board report'
            $options
            Convert-OfficeDocument -Document $document -To Html -AsString -Options $options -ErrorAction Stop
            """);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(2, output.Count);
        var options = Assert.IsType<PdfToHtmlOptions>(output[0].BaseObject);
        Assert.Equal(PdfHtmlProfile.Semantic, options.Profile);
        string html = Assert.IsType<string>(output[1].BaseObject);
        Assert.Contains("Quarterly report content", html, StringComparison.Ordinal);
        Assert.Equal(!fragment, html.Contains("<!doctype html>", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(!fragment, html.Contains("<title>Board report</title>", StringComparison.Ordinal));
        Assert.Equal(!hideMetadata, html.Contains("Report author", StringComparison.Ordinal));
    }
}
