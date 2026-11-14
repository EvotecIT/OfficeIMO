using OfficeIMO.Diff;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfTextRedlineTests {
    private const string First = "The quarterly report covers revenue, costs and hiring across all regions.";
    private const string Second = "Northern offices grew faster than expected after the spring campaign.";
    private const string Third = "Budget requests for the next year are due before the end of October.";

    [Fact]
    public void UnchangedDocumentsProduceNoChangesOrAnnotations() {
        byte[] oldPdf = Build(First, Second, Third);
        byte[] newPdf = Build(First, Second, Third);

        PdfTextRedlineResult result = PdfTextRedline.Compare(oldPdf, newPdf);

        Assert.False(result.HasChanges);
        Assert.Equal(0, result.AnnotationCount);
        Assert.Equal(newPdf, result.Pdf);
        Assert.Empty(Annotations(result.Pdf));
        Assert.Contains("\"total\": 0", result.ToJson(), StringComparison.Ordinal);
    }

    [Fact]
    public void ReplacedWordWithinALineIsHighlightedWithTheOldTextInItsNote() {
        byte[] oldPdf = Build(First, Second, Third);
        byte[] newPdf = Build(First, Second.Replace("faster", "slower"), Third);

        PdfTextRedlineResult result = PdfTextRedline.Compare(oldPdf, newPdf, new PdfTextRedlineOptions { OldName = "v1.pdf", NewName = "v2.pdf" });

        OfficeChange change = Assert.Single(result.Changes.Changes);
        Assert.Equal(OfficeChangeKind.Replaced, change.Kind);
        Assert.Equal("faster", change.OldText);
        Assert.Equal("slower", change.NewText);
        Assert.Equal(1, change.New!.PageNumber);
        OfficeChangeRegion region = Assert.Single(change.New.Regions);
        Assert.True(region.Right > region.Left && region.Top > region.Bottom);

        PdfAnnotation highlight = Assert.Single(Annotations(result.Pdf), static annotation => annotation.Subtype == "Highlight");
        Assert.Equal(1, highlight.PageNumber);
        Assert.Contains("Replaced: faster", highlight.Contents, StringComparison.Ordinal);
        Assert.Equal("OfficeIMO", highlight.Title);
        Assert.Equal(8, highlight.QuadPoints.Count);
        Assert.True(highlight.HasNormalAppearance);
        Assert.Contains(Annotations(result.Pdf), static annotation => annotation.Subtype == "Popup");

        string json = result.ToJson();
        Assert.Contains("\"oldSource\": \"v1.pdf\"", json, StringComparison.Ordinal);
        Assert.Contains("\"kind\": \"replaced\"", json, StringComparison.Ordinal);
        Assert.Contains("\"regions\": [", json, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertionsAreHighlightedAndDeletionsGetACaretCarryingTheOldText() {
        byte[] oldPdf = Build(First, Second, Third);
        byte[] newPdf = Build(First.Replace("costs and hiring", "hiring"), Second.Replace("the spring", "the long spring"), Third);

        PdfTextRedlineResult result = PdfTextRedline.Compare(oldPdf, newPdf, new PdfTextRedlineOptions { CreatePopups = false });

        Assert.Contains(result.Changes.Changes, static change => change.Kind == OfficeChangeKind.Deleted && change.OldText == "costs and");
        Assert.Contains(result.Changes.Changes, static change => change.Kind == OfficeChangeKind.Inserted && change.NewText == "long");
        IReadOnlyList<PdfAnnotation> annotations = Annotations(result.Pdf);
        PdfAnnotation caret = Assert.Single(annotations, static annotation => annotation.Subtype == "Caret");
        Assert.Equal("Deleted: costs and", caret.Contents);
        PdfAnnotation highlight = Assert.Single(annotations, static annotation => annotation.Subtype == "Highlight");
        Assert.Equal("Inserted: long", highlight.Contents);
        Assert.DoesNotContain(annotations, static annotation => annotation.Subtype == "Popup");
        Assert.Equal(2, result.AnnotationCount);
    }

    [Fact]
    public void MovedPageBreaksAreNotChangesAndEditsAcrossPagesKeepTheirPages() {
        byte[] oldPdf = Build(First, Second, null, Third);
        byte[] newPdf = Build(First, null, Second, Third.Replace("October", "November"));

        PdfTextRedlineResult result = PdfTextRedline.Compare(oldPdf, newPdf);

        OfficeChange change = Assert.Single(result.Changes.Changes);
        Assert.Equal(OfficeChangeKind.Replaced, change.Kind);
        Assert.Equal("October", change.OldText);
        Assert.Equal("November", change.NewText);
        Assert.Equal(2, change.Old!.PageNumber);
        Assert.Equal(2, change.New!.PageNumber);
        PdfAnnotation highlight = Assert.Single(Annotations(result.Pdf), static annotation => annotation.Subtype == "Highlight");
        Assert.Equal(2, highlight.PageNumber);
    }

    [Fact]
    public void ReplacementSpanningAPageBreakIsHighlightedOnBothPages() {
        byte[] oldPdf = Build(First, null, Second);
        byte[] newPdf = Build(First.Replace("all regions.", "every region in"), null, Second.Replace("Northern offices", "the north"));

        PdfTextRedlineResult result = PdfTextRedline.Compare(oldPdf, newPdf, new PdfTextRedlineOptions { CreatePopups = false });

        OfficeChange change = Assert.Single(result.Changes.Changes);
        Assert.Equal(new[] { 1, 2 }, change.New!.Regions.Select(static region => region.PageNumber).Distinct());
        IReadOnlyList<PdfAnnotation> highlights = Annotations(result.Pdf).Where(static annotation => annotation.Subtype == "Highlight").ToList();
        Assert.Equal(new int?[] { 1, 2 }, highlights.Select(static annotation => annotation.PageNumber).OrderBy(static page => page));
        Assert.Contains(highlights, static annotation => annotation.Contents!.EndsWith("(continued)", StringComparison.Ordinal));
    }

    [Fact]
    public void ReorderedParagraphsAreReportedAsMoves() {
        byte[] oldPdf = Build(First, Second, Third);
        byte[] newPdf = Build(Second, First, Third);

        PdfTextRedlineResult result = PdfTextRedline.Compare(oldPdf, newPdf);

        OfficeChange moved = Assert.Single(result.Changes.Changes);
        Assert.Equal(OfficeChangeKind.Moved, moved.Kind);
        Assert.True(moved.OldText == First || moved.OldText == Second);
        Assert.NotEqual(moved.Old!.Start, moved.New!.Start);
        PdfAnnotation highlight = Assert.Single(Annotations(result.Pdf), static annotation => annotation.Subtype == "Highlight");
        Assert.StartsWith("Moved from page 1: ", highlight.Contents, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfWithoutATextLayerIsRefused() {
        byte[] withText = Build(First);
        byte[] blank = PdfDocument.Create().ToBytes();

        ArgumentException newError = Assert.Throws<ArgumentException>(() => PdfTextRedline.Compare(withText, blank));
        Assert.Equal("newPdf", newError.ParamName);
        Assert.Contains("no extractable text", newError.Message, StringComparison.Ordinal);
        ArgumentException oldError = Assert.Throws<ArgumentException>(() => PdfTextRedline.Compare(blank, withText));
        Assert.Equal("oldPdf", oldError.ParamName);
    }

    [Fact]
    public void AnnotationLimitKeepsEveryChangeInTheList() {
        byte[] oldPdf = Build(First, Second, Third);
        byte[] newPdf = Build(First.Replace("revenue", "income"), Second.Replace("faster", "slower"), Third.Replace("October", "November"));

        PdfTextRedlineResult result = PdfTextRedline.Compare(oldPdf, newPdf, new PdfTextRedlineOptions { MaxAnnotatedChanges = 1 });

        Assert.Equal(3, result.Changes.Changes.Count);
        Assert.Equal(1, result.AnnotationCount);
        Assert.True(result.AnnotationLimitReached);
    }

    [Fact]
    public void HtmlReportShowsBothTextsAndTheChangeList() {
        byte[] oldPdf = Build(First, Second);
        byte[] newPdf = Build(First, Second.Replace("faster", "slower"));

        string html = PdfTextRedline.Compare(oldPdf, newPdf).ToHtml("Quarterly report");

        Assert.Contains("<title>Quarterly report</title>", html, StringComparison.Ordinal);
        Assert.Contains("<del>faster</del>", html, StringComparison.Ordinal);
        Assert.Contains("<ins>slower</ins>", html, StringComparison.Ordinal);
        Assert.Contains("<h2>Changes</h2>", html, StringComparison.Ordinal);
    }

    private static byte[] Build(params string?[] paragraphs) {
        PdfDocument document = PdfDocument.Create();
        foreach (string? paragraph in paragraphs) {
            if (paragraph is null) document.PageBreak();
            else document.Paragraph(builder => builder.AddText(paragraph));
        }

        return document.ToBytes();
    }

    private static IReadOnlyList<PdfAnnotation> Annotations(byte[] pdf) => PdfDocument.Load(pdf).Inspect().Annotations;
}
