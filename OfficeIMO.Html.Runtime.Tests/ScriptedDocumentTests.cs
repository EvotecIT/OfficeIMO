using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ScriptedDocumentTests {
    [Fact]
    public async Task ExternalProvidersCanReturnIndependentFrozenCaptures() {
        IHtmlScriptRuntimeProvider provider = new ExampleProvider();
        var capture = await provider.CaptureTrustedAsync(new HtmlScriptRequest());
        Assert.True(capture.Document.IsReadOnly);
        Assert.Equal("ExampleProvider/1", capture.ProviderId);
        var mutable = new HtmlDocument(AngleSharpDomServices.Instance, "example");
        Assert.Throws<ArgumentException>(() => new HtmlScriptCapture(mutable, "example"));
    }

    private sealed class ExampleProvider : IHtmlScriptRuntimeProvider {
        public Task<HtmlScriptCapture> CaptureTrustedAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default) =>
            Task.FromResult(new HtmlScriptCapture(new HtmlDocument(AngleSharpDomServices.Instance, "example").Freeze(), "ExampleProvider/1"));
    }

    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task RunsInlineEventsPromisesTimersAndSuppliedScriptsBeforeCapture() {
        var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = """
                <!doctype html><button id="button">Run</button><p id="status">Loading</p>
                <script>
                  document.addEventListener('DOMContentLoaded', () => document.body.setAttribute('data-loaded', 'yes'));
                  document.querySelector('#button').addEventListener('click', () => {
                    Promise.resolve().then(() => setTimeout(() => {
                      document.querySelector('#status').textContent = 'Zażółć € — ready';
                      window.ready = true;
                    }, 20));
                  });
                </script>
                """,
            Scripts = new[] { "document.querySelector('#button').click();" },
            ReadyExpression = "window.ready === true"
        });
        Assert.Equal("yes", capture.Document.Body!.GetAttribute("data-loaded"));
        Assert.Equal("Zażółć € — ready", capture.Document.QuerySelector("#status")!.TextContent);
        var conversion = HtmlConversionDocument.FromDocument(capture.Document);
        Assert.Contains("Zażółć", conversion.ToMarkdown());
        Assert.Contains("ready", PdfReadDocument.Open(conversion.ToPdfBytes()).ExtractText());
        Assert.True(capture.Document.IsReadOnly);
        Assert.Equal(HtmlDocumentMode.Standards, capture.Document.Mode);
        Assert.Contains("AngleSharp.Js/", capture.ProviderId);
        HtmlDocument edited = capture.Document.Edit(doc => doc.QuerySelector("#status")!.TextContent = "Edited");
        Assert.Equal("Edited", edited.QuerySelector("#status")!.TextContent);
        Assert.Equal("Zażółć € — ready", capture.Document.QuerySelector("#status")!.TextContent);
    }

    [Fact]
    public async Task PreservesMutatedStructureNamespacesAndTemplateContentsWithoutHtmlReparse() {
        var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<p id='outer'>before</p><template id='template'><b>original</b></template><svg id='svg'></svg>",
            Scripts = new[] { """
                const nested = document.createElement('p'); nested.textContent = 'nested';
                document.querySelector('#outer').appendChild(nested);
                document.querySelector('#template').content.firstChild.textContent = 'changed';
                const svg = document.querySelector('#svg');
                svg.setAttributeNS('urn:example', 'e:value', 'retained');
                """ }
        });
        HtmlElement outer = capture.Document.QuerySelector("#outer")!;
        Assert.Equal("p", Assert.Single(outer.Children).LocalName);
        Assert.Equal("before", outer.ChildNodes[0].TextContent);
        Assert.Equal("changed", capture.Document.QuerySelector("#template")!.TemplateContent!.TextContent);
        Assert.Equal("retained", capture.Document.QuerySelector("#svg")!.GetAttribute("urn:example", "value"));
        Assert.Null(outer.SourceIndex);
    }

    [Theory]
    [InlineData("inline")]
    [InlineData("supplied")]
    [InlineData("readiness")]
    public async Task TerminatesRunawayExecutionAtTheDeadline(string stage) {
        var request = new HtmlScriptRequest { Html = "<p>hello</p>", Timeout = TimeSpan.FromSeconds(1) };
        if (stage == "inline") request.Html += "<script>while(true){}</script>";
        if (stage == "supplied") request.Scripts = new[] { "while(true){}" };
        if (stage == "readiness") request.ReadyExpression = "(()=>{while(true){}})()";
        await Assert.ThrowsAsync<TimeoutException>(() => Runtime().CaptureTrustedAsync(request));
    }

    [Fact]
    public async Task PreservesDistinctCaseSensitiveAttributesThroughCaptureEditingAndImport() {
        var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<div id='target'></div>",
            Scripts = new[] { "const e=document.querySelector('#target');e.setAttributeNS(null,'DATA-X','upper');e.setAttributeNS(null,'data-x','lower');" },
            ReadyExpression = "document.querySelector('#target').getAttributeNS(null,'DATA-X')==='upper' && document.querySelector('#target').getAttributeNS(null,'data-x')==='lower'"
        });
        static void Check(HtmlElement element) {
            Assert.Equal("upper", element.GetAttribute(string.Empty, "DATA-X"));
            Assert.Equal("lower", element.GetAttribute(string.Empty, "data-x"));
            Assert.Contains("DATA-X=\"upper\"", element.OuterHtml);
            Assert.Contains("data-x=\"lower\"", element.OuterHtml);
        }
        Check(capture.Document.QuerySelector("#target")!);
        var edited = capture.Document.Edit(doc => doc.Body!.SetAttribute("data-edited", "yes"));
        Check(edited.QuerySelector("#target")!);
        var destination = new HtmlDocument(AngleSharpDomServices.Instance, "import");
        var imported = (HtmlElement)destination.ImportNode(edited.QuerySelector("#target")!, deep: true);
        destination.AppendChild(imported);
        Check(imported);
    }

    [Fact]
    public async Task CancellationStopsActiveExecutionAndPreservesCancellationIdentity() {
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(600));
        var error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<script>while(true){}</script>", Timeout = TimeSpan.FromSeconds(20)
        }, cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
    }

    [Fact]
    public async Task FailsInsteadOfPublishingAResultAfterAnInlineScriptError() {
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<p>unchanged</p><script>throw new Error('intentional script failure')</script>"
        }));
        Assert.Contains("intentional script failure", error.Message);
    }

    [Theory]
    [InlineData("nodes")]
    [InlineData("depth")]
    [InlineData("output")]
    public async Task AppliesCaptureBudgetsToScriptGeneratedContent(string budget) {
        var request = new HtmlScriptRequest {
            Html = "<div id='root'></div>",
            Scripts = new[] { "document.querySelector('#root').innerHTML='<section><b>'+ 'x'.repeat(1000) + '</b></section>';" }
        };
        if (budget == "nodes") request.MaxNodes = 4;
        if (budget == "depth") request.MaxDepth = 2;
        if (budget == "output") request.MaxOutputCharacters = 500;
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(request));
    }

    [Fact]
    public async Task ReadinessRequiresBooleanTrueAndInputIsBoundedBeforeExecution() {
        await Assert.ThrowsAsync<TimeoutException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<p>hello</p>", ReadyExpression = "'true'", Timeout = TimeSpan.FromSeconds(1)
        }));
        await Assert.ThrowsAsync<ArgumentException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<p>hello</p>", Scripts = new[] { new string(' ', 100) }, MaxInputCharacters = 50
        }));
    }
}
