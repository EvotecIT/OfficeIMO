using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRedactionSearchPageTests {
    [Fact]
    public void PageSelectionPrecedesCandidateLimitAndRejectsInvalidPages() {
        PdfDocument document = PdfDocument.Create(compose => {
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(text => text.Text("Private account 123")))));
            compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(text => text.Text("Private account 456")))));
        });
        var options = new PdfRedactionSearchOptions { MaximumCandidates = 1 };
        options.AddLiteral("Private account");
        Assert.Throws<InvalidOperationException>(() => document.Redactions.Search(options));
        options.PageNumbers.Add(2);
        PdfRedactionPlan plan = document.Redactions.Search(options);
        Assert.Equal(2, Assert.Single(plan.Areas).PageNumber);
        options.PageNumbers.Add(3);
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Redactions.Search(options));
    }
}
