using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData(0D, 1)]
    [InlineData(8D, 2)]
    [InlineData(16D, 2)]
    public void FollowingParagraphAfterSpacingStaysOutsideTheReservedNoteArea(double after, int pages) {
        PdfOptions options = Options();
        options.DefaultParagraphStyle = BodyAfterStyle(after);
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 8; index++)
                body.Text($"ReservedNote{index:D2}", style: BodyAfterStyle(0));
        });
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: BodyAfterStyle(0));
        // Resolve the following paragraph's after-space from the effective
        // default, as Word does for a paragraph without direct properties.
        document.Content.Text("FollowingBody");

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(pages, pdf.NumberOfPages);
        Assert.Contains("Reference", pdf.GetPage(1).Text, StringComparison.Ordinal);
        for (int index = 0; index < 8; index++)
            Assert.Contains($"ReservedNote{index:D2}", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("FollowingBody", pdf.GetPage(pages).Text, StringComparison.Ordinal);
        if (pages > 1)
            Assert.DoesNotContain("FollowingBody", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0D, 1)]
    [InlineData(8D, 2)]
    [InlineData(16D, 2)]
    public void ReferencingParagraphAfterSpacingParticipatesInTheInitialNoteReservation(double after, int pages) {
        var document = PdfDocument.Create(Options());
        var note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 9; index++)
                body.Text($"AdmittedNote{index:D2}", style: BodyAfterStyle(0));
        });
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("Reference").WithFootnoteReference(note)
        }), style: BodyAfterStyle(after));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(pages, pdf.NumberOfPages);
        Assert.Contains("Reference", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("AdmittedNote08", pdf.GetPage(pages).Text, StringComparison.Ordinal);
        if (pages > 1)
            Assert.DoesNotContain("AdmittedNote08", pdf.GetPage(1).Text, StringComparison.Ordinal);
        foreach (string token in Enumerable.Range(0, 9).Select(index => $"AdmittedNote{index:D2}"))
            Assert.Single(pdf.GetPages(), page => page.Text.IndexOf(token, StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void OrdinaryFinalParagraphRetainsItsPhysicalPageEndSpacingBehavior() {
        var document = PdfDocument.Create(Options());
        document.Content.Text("Reference", style: BodyAfterStyle(0)).Spacer(80)
            .Text("FollowingBody", style: BodyAfterStyle(16));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("FollowingBody", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("heading", false)]
    [InlineData("heading", true)]
    [InlineData("list", false)]
    [InlineData("list", true)]
    public void HeadingAndListAfterSpacingParticipatesInTheNoteBoundary(string owner, bool introducesReference) {
        string? fontPath = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(fontPath);
        var options = Options().EmbedStandardFont(PdfStandardFont.Helvetica, fontPath!, "Carlito");
        var document = PdfDocument.Create(options);
        int noteLineCount = introducesReference ? 9 : 8;
        var note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < noteLineCount; index++)
                body.Text($"OwnerNote{index:D2}", style: BodyAfterStyle(0));
        });
        if (!introducesReference)
            document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("Reference").WithFootnoteReference(note)
            }), style: BodyAfterStyle(0));

        PdfTextRun run = PdfTextRun.Normal("OwnerBody", fontSize: 8, font: PdfStandardFont.Helvetica);
        if (introducesReference) run = run.WithFootnoteReference(note);
        if (owner == "heading") {
            document.Content.Heading(2, "OwnerBody", paragraph => paragraph.Runs(new[] { run }),
                PdfAlign.Left, null, new PdfHeadingStyle {
                    FontSize = 8, Font = PdfStandardFont.Helvetica, Bold = false,
                    LineSpacing = PdfLineSpacing.Exactly(10), SpacingBefore = 0, SpacingAfter = 8,
                    KeepWithNext = false
                });
        } else if (owner == "list") {
            document.Content.RichNumbered(new[] { PdfListItem.Rich(new[] { run }) }, style: new PdfListStyle {
                FontSize = 8, MarkerFontSize = 8, MarkerFont = PdfStandardFont.Helvetica,
                LineSpacing = PdfLineSpacing.Exactly(10), SpacingBefore = 0, SpacingAfter = 8,
                ItemSpacing = 0, KeepTogether = false, KeepWithNext = false
            });
        } else {
            throw new ArgumentOutOfRangeException(nameof(owner));
        }

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("OwnerBody", pdf.GetPage(introducesReference ? 1 : 2).Text, StringComparison.Ordinal);
        for (int index = 0; index < noteLineCount; index++) {
            string token = $"OwnerNote{index:D2}";
            Assert.Single(pdf.GetPages(), page => page.Text.IndexOf(token, StringComparison.Ordinal) >= 0);
            int expectedPage = introducesReference && index == 8 ? 2 : 1;
            Assert.Contains(token, pdf.GetPage(expectedPage).Text, StringComparison.Ordinal);
        }
    }

    private static PdfParagraphStyle BodyAfterStyle(double after) => new() {
        FontSize = 8,
        LineSpacing = PdfLineSpacing.Exactly(10),
        WidowControl = false,
        SpacingBefore = 0,
        SpacingAfter = after
    };
}
