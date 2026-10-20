using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData("heading")]
    [InlineData("list")]
    public void HeadingAndListReferencesMoveWithTheirFirstNoteAndReserveFollowingBody(string owner) {
        string? fontPath = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(fontPath);
        var options = Options().EmbedStandardFont(PdfStandardFont.Helvetica, fontPath!, "Carlito");
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 9; index++) body.Text("OwnerNote" + index, style: Style());
        });
        document.Content.Text("Host", style: Style()).Spacer(75D);
        var reference = PdfTextRun.Normal("BodyReference", fontSize: 8, font: PdfStandardFont.Helvetica).WithFootnoteReference(note);
        if (owner == "heading") {
            // This is the existing converter-facing rich heading entry used by Word.
            document.Content.Heading(2, "BodyReference", paragraph => paragraph.Runs(new[] { reference }),
                PdfAlign.Left, null, new PdfHeadingStyle {
                    FontSize = 8, Font = PdfStandardFont.Helvetica, Bold = false,
                    LineSpacing = PdfLineSpacing.Exactly(10), SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = false
                });
        } else if (owner == "list") {
            document.Content.RichNumbered(new[] { PdfListItem.Rich(new[] { reference }) }, style: new PdfListStyle {
                FontSize = 8, MarkerFontSize = 8, MarkerFont = PdfStandardFont.Helvetica,
                LineSpacing = PdfLineSpacing.Exactly(10), SpacingBefore = 0, SpacingAfter = 0, ItemSpacing = 0,
                KeepTogether = false, KeepWithNext = false
            });
        } else {
            throw new ArgumentOutOfRangeException(nameof(owner));
        }
        document.Content.Text("FollowingBody", style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("Host", pdf.GetPage(1).Text);
        Assert.DoesNotContain("BodyReference", pdf.GetPage(1).Text);
        Assert.DoesNotContain("OwnerNote", pdf.GetPage(1).Text);
        Assert.Contains("BodyReference", pdf.GetPage(2).Text);
        Assert.Contains("OwnerNote0", pdf.GetPage(2).Text);
        Assert.Contains("OwnerNote8", pdf.GetPage(2).Text);
        Assert.DoesNotContain("FollowingBody", pdf.GetPage(2).Text);
        Assert.Contains("FollowingBody", pdf.GetPage(3).Text);
        string paintedText = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        foreach (string marker in new[] { "Host", "BodyReference", "FollowingBody" }.Concat(Enumerable.Range(0, 9).Select(index => "OwnerNote" + index)))
            Assert.Single(System.Text.RegularExpressions.Regex.Matches(paintedText,
                System.Text.RegularExpressions.Regex.Escape(marker)));
    }
}
