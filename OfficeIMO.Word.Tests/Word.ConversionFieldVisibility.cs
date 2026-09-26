using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordConversionFieldVisibilityTests {
    [Fact]
    public void MixedFieldMarkersWithinOneRunKeepOnlyVisibleSegments() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Prefix ");
        paragraph._paragraph.Append(new Run(
            new Text("Before"),
            new FieldChar { FieldCharType = FieldCharValues.Begin },
            new Text("Hidden"),
            new FieldChar { FieldCharType = FieldCharValues.Separate },
            new Text("Shown"),
            new FieldChar { FieldCharType = FieldCharValues.End },
            new Text("After")));

        Assert.Equal("Prefix BeforeShownAfter", string.Concat(paragraph.GetRuns().Select(run => run.Text)));
        Assert.DoesNotContain("Hidden", document.ToHtml(), StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void MixedMarkerRunRetainsSourcePositionAndMutationTarget() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Prefix ");
        var mixed = new Run(
            new FieldChar { FieldCharType = FieldCharValues.Begin },
            new Text("Hidden"),
            new FieldChar { FieldCharType = FieldCharValues.Separate },
            new Text("Before"),
            new FieldChar { FieldCharType = FieldCharValues.End });
        paragraph._paragraph.Append(mixed,
            new M.OfficeMath(new M.Run(new M.Text("Equation"))),
            new Run(new Text("After")));

        WordParagraph projected = Assert.Single(paragraph.GetRuns(), run => ReferenceEquals(run._run, mixed));
        Assert.Equal("Before", projected.Text);
        string html = document.ToHtml();
        Assert.True(html.IndexOf("Before", StringComparison.Ordinal) < html.IndexOf("Equation", StringComparison.Ordinal));
        Assert.True(html.IndexOf("Equation", StringComparison.Ordinal) < html.IndexOf("After", StringComparison.Ordinal));

        projected.Text = "Changed";
        Assert.Equal("Changed", projected.Text);
        Assert.Contains("Changed", mixed.InnerText, StringComparison.Ordinal);
    }

    [Fact]
    public void ComplexFieldInstructionAcrossParagraphsStaysHidden() {
        using WordDocument document = WordDocument.Create();
        WordParagraph start = document.AddParagraph("Start ");
        start._paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        WordParagraph instruction = document.AddParagraph("Hidden instruction");
        WordParagraph result = document.AddParagraph("Result ");
        result._paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("Visible")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        Assert.Empty(instruction.GetRuns());
        Assert.DoesNotContain("Hidden instruction", document.ToHtml(), StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden instruction", document.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains("Visible", document.ToHtml(), StringComparison.Ordinal);
        Assert.Contains("Visible", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void TextBoxFieldMarkersDoNotHideLaterBodyParagraphs() {
        using WordDocument document = WordDocument.Create();
        WordTextBox textBox = document.AddTextBox("Box instruction");
        Paragraph boxParagraph = textBox.Content!.Descendants<Paragraph>().First();
        boxParagraph.InsertAt(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }), 0);
        WordParagraph body = document.AddParagraph("Visible body");

        Assert.Equal("Visible body", string.Concat(body.GetRuns().Select(run => run.Text)));
        Assert.Contains("Visible body", document.ToHtml(), StringComparison.Ordinal);
    }

    [Fact]
    public void SimpleFieldInsideComplexInstructionIsHiddenAcrossRunConverters() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Prefix ");
        paragraph._paragraph.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new SimpleField(new Run(new Text("Hidden"))) { Instruction = " PAGE " },
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new SimpleField(new Run(new Text("Visible"))) { Instruction = " PAGE " },
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        Assert.Equal("Prefix Visible", string.Concat(paragraph.GetRuns().Select(run => run.Text)));
        Assert.DoesNotContain("Hidden", document.ToHtml(), StringComparison.Ordinal);
        Assert.Contains("Visible", document.ToHtml(), StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden", document.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains("Visible", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void HyperlinkAroundContentControlRetainsLinkAndOnlyItsOwnText() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Prefix ");
        WordParagraph linked = paragraph.AddHyperLink("Original", new Uri("https://example.com/"));
        Hyperlink hyperlink = linked._hyperlink!;
        hyperlink.RemoveAllChildren();
        hyperlink.Append(
            new SdtRun(new SdtContentRun(new Run(new Text("Linked")))),
            new Run(new Text(" tail")));

        WordParagraph[] runs = paragraph.GetRuns().ToArray();
        Assert.Equal(new[] { "Prefix ", "Linked", " tail" }, runs.Select(run => run.Text));
        Assert.All(runs.Skip(1), run => Assert.Equal("https://example.com/", run.Hyperlink?.Uri?.ToString()));
        Assert.Contains("href=\"https://example.com/\"", document.ToHtml(), StringComparison.Ordinal);
        Assert.Contains("Linked", document.ToHtml(), StringComparison.Ordinal);
        Assert.Contains("[Linked]", document.ToMarkdown(), StringComparison.Ordinal);
    }
}
