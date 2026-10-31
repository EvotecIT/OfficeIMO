using System.Linq;
using OfficeIMO.Word;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_IndentedHeadingLinesUpWithIndentedBody() {
        using WordDocument document = WordDocument.Create();
        var heading = document.AddParagraph("Headline");
        heading.Style = WordParagraphStyles.Heading1;
        heading.IndentationBefore = OfficeLength.Points((720) / 20d);
        var body = document.AddParagraph("Bodytext");
        body.IndentationBefore = OfficeLength.Points((720) / 20d);

        using PdfPigDocument pdf = OpenAuditPdf(document);
        var words = pdf.GetPage(1).GetWords().ToList();
        double headingLeft = words.First(word => word.Text == "Headline").BoundingBox.Left;
        double bodyLeft = words.First(word => word.Text == "Bodytext").BoundingBox.Left;
        Assert.InRange(headingLeft - bodyLeft, -2D, 2D);
    }
}
