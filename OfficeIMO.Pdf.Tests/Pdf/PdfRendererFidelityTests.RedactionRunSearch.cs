using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRendererFidelityTests {
    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(270)]
    public void KernedPhraseSearchProducesOnePreciseAreaPerOccurrence(int rotation) {
        PdfDocument document = PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800)
            .Content(content => content.Item(item => item.Paragraph(text => text.Text("Before private account 123 and private account 456 after"))))));
        if (rotation != 0) document = document.Pages.Rotate(rotation, 1);
        var search = new PdfRedactionSearchOptions { TextSelection = PdfRedactionTextSelection.MatchedGlyphs };
        search.AddLiteral("private account");
        PdfRedactionPlan plan = document.Redactions.Search(search);
        Assert.True(plan.IsReviewable, string.Join(";", plan.Findings.Select(finding => finding.Message)));
        Assert.Equal(2, plan.Areas.Count);
    }
}
