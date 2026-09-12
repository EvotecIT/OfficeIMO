using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeFormStateTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task LiveValuesAndDefaultsSurviveCaptureCloneImportAndConversionAfterDisposal() {
        var captured = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = """
                <!doctype html><input id="text" name="text" value="Original">
                <textarea id="area" name="area">Original area</textarea>
                <input id="check" name="check" type="checkbox" checked>
                <select id="choice" name="choice" multiple><option selected value="a">Alpha</option><option value="b">Beta</option></select>
                <select id="empty" name="empty"><option selected>Default</option></select>
                """,
            Scripts = new[] { """
                document.querySelector('#text').value='Live input';
                document.querySelector('#area').value='Live area';
                document.querySelector('#check').checked=false;
                document.querySelector('#check').indeterminate=true;
                document.querySelector('#choice').options[0].selected=false;
                document.querySelector('#choice').options[1].selected=true;
                document.querySelector('#empty').options[0].selected=false;
                """ }
        });
        HtmlDocument source = captured.Document;
        Assert.True(source.IsReadOnly);
        Assert.Throws<InvalidOperationException>(() => source.QuerySelector("#text")!.FormState = null);
        var imported = new HtmlDocument(AngleSharpDomServices.Instance, "test");
        imported.AppendChild(imported.ImportNode(source.DocumentElement!));
        foreach (HtmlDocument document in new[] { source, source.Clone(), source.CloneAttached(), source.Edit(_ => { }), imported }) {
            Assert.Equal("Original", document.QuerySelector("#text")!.GetAttribute("value"));
            Assert.Equal("Original area", document.QuerySelector("#area")!.TextContent);
            Assert.Contains("Original area</textarea>", document.OuterHtml);
            Assert.DoesNotContain("Live input", document.OuterHtml);
            Assert.Equal("Live input", document.QuerySelector("#text")!.FormState!.Value);
            Assert.False(document.QuerySelector("#check")!.FormState!.IsChecked);
            Assert.True(document.QuerySelector("#check")!.FormState!.IsIndeterminate);
            Assert.Single(document.QuerySelectorAll("#choice option:checked"));
            Assert.Equal("b", document.QuerySelector("#choice option:checked")!.GetAttribute("value"));
            var conversion = HtmlConversionDocument.FromDocument(document);
            var controls = Flatten(conversion.SemanticDocument.Sections.SelectMany(section => section.Blocks))
                .Where(block => block.FormControl != null).Select(block => block.FormControl!).ToDictionary(control => control.Name);
            Assert.Equal("Live input", controls["text"].Value);
            Assert.Equal("Live area", controls["area"].Value);
            Assert.False(controls["check"].IsChecked);
            Assert.True(controls["check"].IsIndeterminate);
            Assert.Equal(new[] { "b" }, controls["choice"].Values);
            Assert.Empty(controls["empty"].Values);
            var normalized = conversion.CreateDocumentForConversion();
            Assert.Equal("Live input", normalized.QuerySelector("#text")!.FormState!.Value);
            Assert.Equal("Original", normalized.QuerySelector("#text")!.GetAttribute("value"));
            var render = HtmlRenderEngine.Render(conversion);
            Assert.Contains("Live input", render.Text);
            Assert.Contains("Live area", render.Text);
            Assert.Contains("Beta", render.Text);
            Assert.DoesNotContain("Original", render.Text);
        }
        string pdfText = PdfReadDocument.Open(HtmlConversionDocument.FromDocument(source).ToPdfBytes()).ExtractText();
        Assert.Contains("Live input", pdfText);
        Assert.Contains("Live area", pdfText);
        var pdfResult = HtmlConversionDocument.FromDocument(source).ToPdfDocumentResult();
        Assert.Contains(pdfResult.Warnings, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FormFieldIndeterminateStaticFallback);
        var edited = source.Edit(document => document.QuerySelector("#text")!.FormState = new(HtmlFormControlStateKind.Input, "Edited"));
        Assert.Equal("Live input", source.QuerySelector("#text")!.FormState!.Value);
        Assert.Equal("Edited", edited.QuerySelector("#text")!.FormState!.Value);
        var reset = edited.Edit(document => document.QuerySelector("#text")!.FormState = null);
        Assert.Contains("Original", HtmlRenderEngine.Render(HtmlConversionDocument.FromDocument(reset)).Text);
    }

    [Fact]
    public async Task LiveValuesCountAgainstTheCaptureDataBudget() {
        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<input>", Scripts = new[] { "document.querySelector('input').value='x'.repeat(4096)" }, MaxOutputCharacters = 1024
        }));
        Assert.Contains("budget", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OwnedStateRejectsWrongElementsAndCountsAgainstConversionLimits() {
        var document = HtmlConversionDocument.Parse("<input><textarea></textarea><svg><input/></svg>").Document.Clone();
        var input = document.QuerySelector("input")!;
        Assert.Throws<ArgumentException>(() => input.FormState = new(HtmlFormControlStateKind.Select));
        Assert.Throws<ArgumentNullException>(() => new HtmlFormControlState(HtmlFormControlStateKind.Input));
        Assert.Throws<ArgumentException>(() => new HtmlFormControlState(HtmlFormControlStateKind.Option, "value"));
        Assert.Throws<ArgumentException>(() => new HtmlFormControlState(HtmlFormControlStateKind.TextArea, "value", isChecked: true));
        input.FormState = new(HtmlFormControlStateKind.Input, new string('x', 2048));
        var error = Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.FromDocument(document,
            new() { Limits = new() { MaxInputCharacters = 1024 } }));
        Assert.Equal("MaxInputCharacters", error.LimitSource);
    }

    [Fact]
    public async Task EditedSelectConstraintsAndImportedOptionsRetainEffectiveSelection() {
        var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
            Html = "<select name='choice' multiple><option value='a' selected>Alpha</option><option value='b'>Beta</option></select>",
            Scripts = new[] { "document.querySelectorAll('option')[1].selected=true" }
        });
        var dropdown = capture.Document.Edit(document => document.QuerySelector("select")!.RemoveAttribute("multiple"));
        Assert.Equal("b", Assert.Single(dropdown.QuerySelectorAll("option:checked")).GetAttribute("value"));
        Assert.Equal(new[] { "b" }, Selection(dropdown));
        Assert.Contains("Beta", HtmlRenderEngine.Render(HtmlConversionDocument.FromDocument(dropdown)).Text);

        var destination = HtmlConversionDocument.Parse("<select name='choice' multiple><option value='original'>Original</option></select>").Document.Clone();
        destination.QuerySelector("select")!.AppendChild(destination.ImportNode(capture.Document.QuerySelectorAll("option")[1]));
        Assert.Equal(new[] { "b" }, Selection(destination));
        destination.QuerySelector("option")!.FormState = new(HtmlFormControlStateKind.Option, isSelected: true);
        Assert.Equal(new[] { "original", "b" }, Selection(destination));
        foreach (var option in destination.QuerySelectorAll("option")) option.FormState = new(HtmlFormControlStateKind.Option);
        destination.QuerySelector("select")!.RemoveAttribute("multiple");
        Assert.Empty(Selection(destination));

        var partial = HtmlConversionDocument.Parse("<select name='choice'><option value='a'>Alpha</option><option value='b'>Beta</option></select>").Document.Clone();
        partial.QuerySelectorAll("option")[1].FormState = new(HtmlFormControlStateKind.Option);
        Assert.Equal(new[] { "a" }, Selection(partial));
        partial.QuerySelector("select")!.AppendChild(partial.ImportNode(destination.QuerySelector("option")!));
        Assert.Equal(new[] { "a" }, Selection(partial));
        Assert.Contains("Alpha", HtmlRenderEngine.Render(HtmlConversionDocument.FromDocument(partial)).Text);
        partial.QuerySelectorAll("option")[0].FormState = new(HtmlFormControlStateKind.Option);
        Assert.Empty(Selection(partial));
    }

    [Fact]
    public async Task RenderBudgetsIncludeLiveValuesAcrossSyncAsyncAndImageEntryPoints() {
        var document = HtmlConversionDocument.Parse("<textarea>Default</textarea>").Document.Edit(tree =>
            tree.QuerySelector("textarea")!.FormState = new(HtmlFormControlStateKind.TextArea, new string('x', 4096)));
        var conversion = HtmlConversionDocument.FromDocument(document);
        var options = new HtmlRenderOptions { MaxInputCharacters = 1024 };
        Assert.Throws<HtmlDomLimitException>(() => HtmlRenderEngine.Render(conversion, options));
        await Assert.ThrowsAsync<HtmlDomLimitException>(() => HtmlRenderEngine.RenderAsync(conversion, options));
        await Assert.ThrowsAsync<HtmlDomLimitException>(() => conversion.ToPngAsync(options));
        Assert.Throws<HtmlDomLimitException>(() => conversion.ToPdfBytes(new() { MaxInputCharacters = 1024 }));
        await Assert.ThrowsAsync<HtmlDomLimitException>(() => conversion.ToPdfBytesAsync(new() { MaxInputCharacters = 1024 }));
        var combined = new HtmlRenderOptions { MaxInputCharacters = 4096 };
        Assert.Throws<HtmlDomLimitException>(() => HtmlRenderEngine.Render(conversion, combined));
    }

    private static IReadOnlyList<string> Selection(HtmlDocument document) =>
        Flatten(HtmlConversionDocument.FromDocument(document).SemanticDocument.Sections.SelectMany(section => section.Blocks))
            .Single(block => block.FormControl?.Name == "choice").FormControl!.Values;

    private static IEnumerable<HtmlSemanticBlock> Flatten(IEnumerable<HtmlSemanticBlock> blocks) {
        foreach (var block in blocks) {
            yield return block;
            foreach (var child in Flatten(block.Children)) yield return child;
        }
    }
}
