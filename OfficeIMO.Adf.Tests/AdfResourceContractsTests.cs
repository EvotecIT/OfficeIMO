using System;
using System.IO;
using System.Threading;
using OfficeIMO.Adf;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfResourceContractsTests {
    [Fact]
    public void MutableMarkdownInputIsCheckedBeforeRecursiveConversion() {
        var quote = new QuoteBlock(Array.Empty<string>());
        var document = MarkdownDoc.Create().Add(quote);
        quote.ChildBlocks.Add(quote);
        Assert.Throws<InvalidOperationException>(() => AdfConverter.FromMarkdown(document));

        var normal = MarkdownDoc.Create().P("text");
        Assert.Throws<InvalidDataException>(() => AdfConverter.FromMarkdown(normal, new AdfConversionOptions { MaxNodes = 1 }));
        Assert.Throws<InvalidDataException>(() => AdfConverter.FromMarkdown(normal, new AdfConversionOptions { MaxDepth = 1 }));
    }

    [Fact]
    public void JsonInputBudgetUsesUtf8Bytes() {
        const string json = "{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"paragraph\",\"content\":[{\"type\":\"text\",\"text\":\"🙂\"}]}]}";
        Assert.Throws<InvalidDataException>(() => AdfDocument.Parse(json, new AdfProcessingOptions { MaxInputBytes = json.Length }));
        Assert.Single(AdfDocument.Parse(json, new AdfProcessingOptions { MaxInputBytes = System.Text.Encoding.UTF8.GetByteCount(json) }).Content);
    }

    [Fact]
    public void NativeAndProjectionOperationsRejectCyclesWithoutRecursing() {
        var paragraph = new AdfNode("paragraph");
        paragraph.Content.Add(paragraph);
        var document = new AdfDocument(new[] { paragraph });
        Assert.Throws<InvalidOperationException>(() => document.ToJson());
        Assert.Throws<InvalidOperationException>(() => document.Validate());
        Assert.Throws<InvalidOperationException>(() => AdfConverter.ToMarkdown(document));
    }

    [Fact]
    public void RepeatedReferencesAreAllowedAndConsumeTheNodeBudget() {
        var paragraph = new AdfNode("paragraph");
        var document = new AdfDocument(new[] { paragraph, paragraph });
        Assert.Contains("paragraph", document.ToJson(new AdfProcessingOptions { MaxNodes = 2 }));
        Assert.Throws<InvalidDataException>(() => document.ToJson(new AdfProcessingOptions { MaxNodes = 1 }));
    }

    [Fact]
    public void DepthAndTextLimitsApplyToTheMutableModel() {
        var outer = new AdfNode("paragraph");
        outer.Content.Add(AdfNode.TextNode("abcd"));
        var document = new AdfDocument(new[] { outer });
        Assert.Throws<InvalidDataException>(() => document.Validate(new AdfProcessingOptions { MaxDepth = 1 }));
        Assert.Throws<InvalidDataException>(() => AdfConverter.ToMarkdown(document, new AdfConversionOptions { MaxTextCharacters = 3 }));
    }

    [Fact]
    public void OutputLimitsRejectRatherThanTruncateResults() {
        var document = new AdfDocument(new[] { new AdfNode("paragraph") { Content = { AdfNode.TextNode("abc") } } });
        Assert.Throws<InvalidDataException>(() => document.ToJson(new AdfProcessingOptions { MaxOutputBytes = 20 }));
        Assert.Throws<InvalidDataException>(() => AdfConverter.ToMarkdown(document, new AdfConversionOptions { MaxOutputCharacters = 2 }));
        Assert.Throws<InvalidDataException>(() => AdfConverter.FromMarkdown("abcd", new AdfConversionOptions { MaxInputBytes = 3 }));
    }

    [Fact]
    public void SemanticFallbackAndUnknownJsonTextShareTheTextBudget() {
        var node = new AdfNode("mention").SetAttribute("text", "long label");
        var document = new AdfDocument(new[] { new AdfNode("paragraph") { Content = { node } } });
        Assert.Throws<InvalidDataException>(() => AdfConverter.ToMarkdown(document, new AdfConversionOptions { MaxTextCharacters = 3 }));
        Assert.Throws<InvalidDataException>(() => document.ToJson(new AdfProcessingOptions { MaxTextCharacters = 3 }));
    }

    [Fact]
    public void CanceledOperationsStopBeforeParsingTraversingOrWriting() {
        using var source = new CancellationTokenSource();
        source.Cancel();
        var limits = new AdfProcessingOptions { CancellationToken = source.Token };
        var projection = new AdfConversionOptions { CancellationToken = source.Token };
        var document = new AdfDocument();
        Assert.Throws<OperationCanceledException>(() => AdfDocument.Parse("{}", limits));
        Assert.Throws<OperationCanceledException>(() => document.Validate(limits));
        Assert.Throws<OperationCanceledException>(() => document.ToJson(limits));
        Assert.Throws<OperationCanceledException>(() => AdfConverter.ToMarkdown(document, projection));
        Assert.Throws<OperationCanceledException>(() => AdfConverter.FromMarkdown("text", projection));
    }
}
