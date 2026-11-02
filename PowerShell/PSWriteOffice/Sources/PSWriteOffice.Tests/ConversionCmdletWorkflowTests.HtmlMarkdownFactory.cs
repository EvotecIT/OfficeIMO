using System.Management.Automation;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void HtmlMarkdownFactoryFeedsNativeConversionAndRetainsCopiedWriterSettings(bool dropUnsupported) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-html-markdown-factory-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var source = HtmlConversionDocument.Parse("<html><head><title>Document chrome</title></head><body>" +
                "<h1>Report</h1><ul><li>One</li><li>Two</li></ul><video src='movie.mp4'><p>Video transcript</p></video><p>Final.</p></body></html>");
            var original = new MarkdownHtmlImportOptions {
                UseBodyContentsOnly = false, PreserveUnsupportedInlineHtml = false,
                MarkdownWriteOptions = new MarkdownSaveOptions { UnorderedListMarker = '*', OutputLineEnding = "\r\n",
                    ImageRenderingMode = MarkdownImageRenderingMode.Html }
            };
            string destination = Path.Combine(root, "report.md"), rejected = Path.Combine(root, "rejected.md");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("original", original);
            runspace.SessionStateProxy.SetVariable("dropUnsupported", dropUnsupported);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            runspace.SessionStateProxy.SetVariable("rejected", rejected);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                $options = New-OfficeHtmlMarkdownOptions -BaseOptions $original -IncludeDocumentChrome:$false `
                    -DropUnsupportedBlocks:$dropUnsupported -DropUnsupportedInlineHtml:$false -LineEnding LF -MaxInputCharacters 4000
                Convert-OfficeDocument -Document $source -To Markdown -OutputPath $destination -Options $options -ErrorAction Stop | Out-Null
                $options
                if ($dropUnsupported) {
                    $limited = New-OfficeHtmlMarkdownOptions -BaseOptions $options -MaxInputCharacters 4
                    $unbounded = New-OfficeHtmlMarkdownOptions -BaseOptions $limited -MaxInputCharacters $null
                    Convert-OfficeDocument -Document $source -To Markdown -AsString -Options $unbounded -ErrorAction Stop
                }
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(dropUnsupported ? 2 : 1, output.Count);
            var options = Assert.IsType<MarkdownHtmlImportOptions>(output[0].BaseObject);
            Assert.NotSame(original, options); Assert.NotSame(original.MarkdownWriteOptions, options.MarkdownWriteOptions);
            Assert.True(options.UseBodyContentsOnly); Assert.True(options.PreserveUnsupportedInlineHtml);
            Assert.Equal(!dropUnsupported, options.PreserveUnsupportedBlocks);
            Assert.True(original.PreserveUnsupportedBlocks); Assert.False(original.PreserveUnsupportedInlineHtml);
            Assert.False(original.UseBodyContentsOnly); Assert.Null(original.MaxInputCharacters);
            Assert.Equal("\r\n", original.MarkdownWriteOptions!.OutputLineEnding);
            Assert.Equal(MarkdownImageRenderingMode.Html, options.MarkdownWriteOptions!.ImageRenderingMode);
            string saved = File.ReadAllText(destination);
            Assert.Contains("# Report", saved, StringComparison.Ordinal);
            Assert.Contains("* One\n* Two", saved, StringComparison.Ordinal);
            Assert.DoesNotContain("\r", saved, StringComparison.Ordinal);
            Assert.DoesNotContain("Document chrome", saved, StringComparison.Ordinal);
            Assert.Contains("Video transcript", saved, StringComparison.Ordinal);
            Assert.Equal(!dropUnsupported, saved.Contains("<video", StringComparison.Ordinal));
            if (dropUnsupported) {
                Assert.Contains("# Report", Assert.IsType<string>(output[1].BaseObject), StringComparison.Ordinal);
                shell.Commands.Clear(); shell.Streams.Error.Clear();
                shell.AddScript("""
                    try {
                        Convert-OfficeDocument -Document $source -To Markdown -OutputPath $rejected -Options $limited -ErrorAction Stop
                        $false
                    } catch { $true }
                    """);
                Assert.True(Assert.IsType<bool>(Assert.Single(shell.Invoke()).BaseObject));
                Assert.False(File.Exists(rejected));
            }
        } finally { Directory.Delete(root, true); }
    }
}
