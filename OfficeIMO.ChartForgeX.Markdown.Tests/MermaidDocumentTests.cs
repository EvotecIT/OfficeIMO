using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using global::ChartForgeX.Diagnostics;
using global::ChartForgeX.Markup;
using OfficeIMO.ChartForgeX.Markdown;
using OfficeIMO.Markdown;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.MarkdownRenderer;
using StaticMarkdownRenderer = OfficeIMO.MarkdownRenderer.MarkdownRenderer;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.Markup.Word;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.ChartForgeX.Markdown.Tests;

public sealed class MermaidDocumentTests {
    private const string Fence = "```mermaid title=\"Approval flow\"\nflowchart LR\nA[Request] --> B[Approved]\n```";

    [Fact]
    public void NestedFencesBecomeImagesWhileCaptionsAndFailedSourceRemain() {
        string source = "# Report\n\n> ```mermaid title=\"Approval flow\"\n> flowchart LR\n> A --> B\n> ```\n\n```csharp\nthrow;\n```\n\n```mermaid\nnot-a-diagram\n```";
        var document = MarkdownReader.Parse(source);
        var fences = document.DescendantObjectsOfType<CodeBlock>().ToArray();
        fences[0].Caption = "Request approval";
        var failed = fences.Single(block => block.Content == "not-a-diagram");
        var diagnostics = new List<MarkupDiagnostic>();

        MermaidMarkdownAdapter.Materialize(document, diagnostics.Add);

        var image = Assert.Single(document.DescendantObjectsOfType<ImageBlock>());
        Assert.Equal("Request approval", image.Caption);
        Assert.Equal("Approval flow", image.PlainAlt);
        Assert.Same(failed, document.DescendantObjectsOfType<CodeBlock>().Single(block => block.Language == "mermaid"));
        Assert.Contains(document.DescendantObjectsOfType<CodeBlock>(), block => block.Language == "csharp");
        var error = Assert.Single(diagnostics, item => item.Severity == VisualDiagnosticSeverity.Error);
        Assert.True(error.Line >= failed.SourceSpan!.Value.StartLine);
    }

    [Fact]
    public void WordAndPdfEmbedTheRenderedImageAndPreserveCaption() {
        var document = MarkdownReader.Parse("# Report\n\n" + Fence);
        document.DescendantObjectsOfType<CodeBlock>().Single().Caption = "Request approval";
        MermaidMarkdownAdapter.Materialize(document);
        var image = Assert.Single(document.DescendantObjectsOfType<ImageBlock>());
        var png = Convert.FromBase64String(image.Path.Substring(image.Path.IndexOf(',') + 1));

        using var word = document.ToWordDocument(new MarkdownToWordOptions { FitImagesToContextWidth = true });
        AssertWordImageFitsContent(word);
        using var stream = new MemoryStream();
        word.Save(stream);
        stream.Position = 0;
        using var package = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var media = Assert.Single(package.Entries, entry => entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase));
        using var mediaStream = media.Open();
        using var copy = new MemoryStream();
        mediaStream.CopyTo(copy);
        Assert.Equal(png, copy.ToArray());
        using var xmlReader = new StreamReader(package.GetEntry("word/document.xml")!.Open());
        var xml = XDocument.Parse(xmlReader.ReadToEnd());
        Assert.Contains("Request approval", xml.Root!.Value);
        Assert.Contains(xml.Descendants().Attributes("descr"), value => value.Value == image.PlainAlt);

        var pdf = document.ToPdfBytes();
        var text = Encoding.ASCII.GetString(pdf);
        Assert.StartsWith("%PDF-", text);
        Assert.Contains("/Subtype /Image", text);
        Assert.DoesNotContain("flowchart LR", text);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void WordMarkupEmbedsTransformedImagesAndHonorsResourcePolicy(bool allowed, bool oversized) {
        var reader = MarkdownReaderOptions.CreateOfficeIMOProfile();
        reader.DocumentTransforms.Add(MermaidMarkdownAdapter.CreateTransform());
        var parsed = OfficeMarkupParser.Parse(Fence + "\n_Request approval_", new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Document,
            MarkdownOptions = reader
        });
        using var word = parsed.Document.ToWordDocument(new MarkupToWordOptions {
            AllowDataUriImages = allowed,
            MaximumDataUriImageBytes = oversized ? 1 : 16L * 1024 * 1024
        });
        if (allowed && !oversized) {
            AssertWordImageFitsContent(word);
            Assert.Contains(word.Paragraphs, paragraph => paragraph.Text == "Request approval");
        } else {
            Assert.Empty(word.Images);
            Assert.Contains(word.Paragraphs, paragraph => paragraph.Text.Contains("Approval flow"));
            Assert.DoesNotContain(word.Paragraphs, paragraph => paragraph.Text.Contains("data:image"));
        }
    }

    [Fact]
    public void StaticHtmlIsDeterministicAndIsolatesRepeatedSvgDocuments() {
        var diagnostics = new List<MarkupDiagnostic>();
        var options = new MarkdownRendererOptions();
        MermaidMarkdownAdapter.ConfigureHtml(options, diagnostics.Add);
        MermaidMarkdownAdapter.ConfigureHtml(options, diagnostics.Add);
        string source = "# Report\n\n" + Fence + "\n_Approval caption_\n\n" + Fence + "\n\n```mermaid\nnot-a-diagram\n```";

        string html = StaticMarkdownRenderer.RenderBodyHtml(source, options);

        Assert.False(options.Mermaid.Enabled);
        var matches = Regex.Matches(html, "src=\"data:image/svg\\+xml;base64,([^\"]+)\"");
        Assert.Equal(2, matches.Count);
        foreach (Match match in matches) {
            string svg = Encoding.UTF8.GetString(Convert.FromBase64String(match.Groups[1].Value));
            var root = XDocument.Parse(svg).Root!;
            Assert.Equal("svg", root.Name.LocalName);
            Assert.DoesNotContain(root.Descendants(), element => element.Name.LocalName == "script");
        }
        Assert.Contains("alt=\"Approval flow\"", html);
        Assert.Contains("<div class=\"caption\">Approval caption</div>", html);
        Assert.Contains("language-mermaid", html);
        Assert.DoesNotContain("<script", html);
        var error = Assert.Single(diagnostics, item => item.Severity == VisualDiagnosticSeverity.Error);
        int invalidLine = source.Substring(0, source.IndexOf("not-a-diagram", StringComparison.Ordinal)).Count(character => character == '\n') + 1;
        Assert.Equal(invalidLine, error.Line);
        Assert.Equal(html, StaticMarkdownRenderer.RenderBodyHtml(source, options));
    }

    [Theory]
    [InlineData("# Report\n\n", "")]
    [InlineData("@slide title=\"Report\"\n\n", "")]
    [InlineData("~~~~officeimo-slide\ntitle=\"Report\"\n\n", "\n~~~~")]
    public void PowerPointUsesTheReaderTransformInsidePresentationAuthoring(string prefix, string suffix) {
        var reader = MarkdownReaderOptions.CreateOfficeIMOProfile();
        reader.DocumentTransforms.Add(MermaidMarkdownAdapter.CreateTransform());
        var parsed = OfficeMarkupParser.Parse(prefix + Fence + "\n_Request approval_" + suffix, new OfficeMarkupParserOptions {
            Profile = OfficeMarkupProfile.Presentation,
            MarkdownOptions = reader
        });
        var image = Assert.Single(parsed.Document.DescendantsAndSelf().OfType<OfficeMarkupImageBlock>());
        Assert.Equal("Request approval", image.Caption);
        using var presentation = parsed.Document.ToPowerPointPresentation();
        var picture = Assert.Single(presentation.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointPicture>());
        Assert.Equal(image.Alt, picture.AltText);
        Assert.Equal("Approval flow", picture.AltText);
        var png = Convert.FromBase64String(image.Source.Substring(image.Source.IndexOf(',') + 1));
        Assert.Equal(png, picture.GetImageBytes());
        Assert.True(OfficeImageReader.TryValidateContent(png, null, out var info));
        Assert.Equal(info.Width / (double)info.Height, picture.WidthInches / picture.HeightInches, precision: 4);
        var caption = Assert.Single(presentation.Slides.SelectMany(slide => slide.TextBoxes), text => text.Text == "Request approval");
        Assert.True(caption.TopInches >= picture.TopInches + picture.HeightInches);
        using var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;
        using var package = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.Single(package.Entries, entry => entry.FullName.StartsWith("ppt/media/", StringComparison.Ordinal));
    }

    [Fact]
    public void PowerPointImageCaptionFollowsExplicitPlacement() {
        var rendered = MermaidMarkdownAdapter.Materialize(MarkdownReader.Parse(Fence))
            .DescendantObjectsOfType<ImageBlock>().Single();
        var markup = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        markup.Blocks.Add(new OfficeMarkupImageBlock(rendered.Path, rendered.PlainAlt) {
            Caption = "Request approval",
            Placement = new OfficeMarkupPlacement { X = "3", Y = "2", Width = "4", Height = "1" }
        });
        using var presentation = markup.ToPowerPointPresentation();
        var caption = Assert.Single(presentation.Slides.SelectMany(slide => slide.TextBoxes), text => text.Text == "Request approval");
        Assert.Equal(3, caption.LeftInches, precision: 4);
        Assert.Equal(3.12, caption.TopInches, precision: 4);
        Assert.Equal(4, caption.WidthInches, precision: 4);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, true)]
    public void PowerPointRejectsDisabledOversizedAndMismatchedEmbeddedImages(bool allowed, bool oversized, bool mismatched) {
        var document = MermaidMarkdownAdapter.Materialize(MarkdownReader.Parse(Fence));
        var rendered = document.DescendantObjectsOfType<ImageBlock>().Single();
        var markup = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        markup.Blocks.Add(new OfficeMarkupImageBlock(mismatched
            ? rendered.Path.Replace("image/png", "image/jpeg") : rendered.Path, "Diagram"));
        using var presentation = markup.ToPowerPointPresentation(new MarkupToPowerPointOptions {
            AllowDataUriImages = allowed,
            MaximumDataUriImageBytes = oversized ? 1 : 16L * 1024 * 1024
        });
        Assert.Empty(presentation.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointPicture>());
        using var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;
        using var package = new ZipArchive(stream, ZipArchiveMode.Read);
        Assert.DoesNotContain(package.Entries, entry => entry.FullName.StartsWith("ppt/media/", StringComparison.Ordinal));
        using var slideReader = new StreamReader(package.Entries.Single(entry => entry.FullName.StartsWith("ppt/slides/", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).Open());
        string slideXml = slideReader.ReadToEnd();
        Assert.Contains("Diagram", slideXml);
        Assert.DoesNotContain("data:image", slideXml);
    }

    [Theory]
    [InlineData("EvotecLogo.png", "image/png")]
    [InlineData("Kulek.jpg", "image/jpeg")]
    public void EmbeddedImagesRequireCompleteMatchingContentWithinTheExactByteBudget(string file, string mime) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Images", file));
        string uri = "data:" + mime + ";base64," + Convert.ToBase64String(bytes);
        Assert.True(OfficeImageReader.TryReadBase64DataUri(uri, bytes.Length, out var decoded, out _));
        Assert.Equal(bytes, decoded);
        Assert.False(OfficeImageReader.TryReadBase64DataUri(uri, bytes.Length - 1, out _, out _));
        string wrongMime = mime == "image/png" ? "image/jpeg" : "image/png";
        Assert.False(OfficeImageReader.TryReadBase64DataUri(uri.Replace(mime, wrongMime), bytes.Length, out _, out _));
        Assert.False(OfficeImageReader.TryReadBase64DataUri("data:" + mime + ";base64," + Convert.ToBase64String(bytes.Take(24).ToArray()), bytes.Length, out _, out _));
        var markup = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        markup.Blocks.Add(new OfficeMarkupImageBlock(uri, "Embedded image"));
        using var presentation = markup.ToPowerPointPresentation();
        Assert.Equal(bytes, Assert.Single(presentation.Slides.SelectMany(slide => slide.Shapes).OfType<PowerPointPicture>()).GetImageBytes());
        markup.Profile = OfficeMarkupProfile.Document;
        using var word = markup.ToWordDocument();
        AssertWordImageFitsContent(word);
    }

    private static void AssertWordImageFitsContent(OfficeIMO.Word.WordDocument document) {
        var image = Assert.Single(document.Images);
        var section = document.Sections.First();
        double available = ((double)section.PageSettings.Width!.Value - (double)section.Margins.Left - (double)section.Margins.Right) / 15;
        Assert.InRange(image.Width!.Value, 1, available + 0.001);
    }

    [Fact]
    public void PowerPointSourceFallbackRetainsTheAuthoredDiagram() {
        var markup = new OfficeMarkupDocument(OfficeMarkupProfile.Presentation);
        markup.Blocks.Add(new OfficeMarkupDiagramBlock("mermaid", "flowchart LR\nA --> B"));
        using var presentation = markup.ToPowerPointPresentation(new MarkupToPowerPointOptions { RenderMermaidDiagrams = false });
        using var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;
        using var package = new ZipArchive(stream, ZipArchiveMode.Read);
        using var reader = new StreamReader(package.Entries.Single(entry => entry.FullName.StartsWith("ppt/slides/", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).Open());
        Assert.Contains("A --> B", XDocument.Parse(reader.ReadToEnd()).Root!.Value);
    }
}
