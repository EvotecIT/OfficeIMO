using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordHyperlinkContentControlScopeTests {
    [Fact]
    public void NestedControlRunKeepsTextAndFormattingWithinItsOwnContent() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Anchor");
        paragraph._run!.Remove();
        var hyperlink = new Hyperlink(
            new Run(new Text("Before ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
            new SdtRun(new SdtProperties(), new SdtContentRun(new Run(new Text("Inner")))),
            new Run(new Text(" After") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })) { Anchor = "target" };
        paragraph._paragraph.Append(hyperlink);

        var runs = paragraph.GetRuns().ToArray();
        Assert.Equal(new[] { "Before ", "Inner", " After" }, runs.Select(run => run.Text));
        runs[1].Text = "Changed";
        runs[1].Bold = true;
        Assert.Equal("Before Changed After", hyperlink.InnerText);
        Assert.Null(hyperlink.Elements<Run>().First().RunProperties?.Bold);
        Assert.NotNull(hyperlink.Descendants<SdtRun>().Single().Descendants<Run>().Single().RunProperties?.Bold);
        Assert.Null(hyperlink.Elements<Run>().Last().RunProperties?.Bold);

        using var package = document.ToStream();
        using var reopened = WordDocument.Load(package);
        Assert.Equal(new[] { "Before ", "Changed", " After" }, reopened.Paragraphs.Single().GetRuns().Select(run => run.Text));
    }
}
