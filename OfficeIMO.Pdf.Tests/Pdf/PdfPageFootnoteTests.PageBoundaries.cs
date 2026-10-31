using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData("Section")]
    [InlineData("PanelSection")]
    [InlineData("PageEntry")]
    [InlineData("PageExit")]
    public void ExplicitSectionAndPageBoundariesFinishNotesBeforeChangingTheBodyFrame(string boundary) {
        PdfCreateOptions options = Options();
        if (boundary == "PageExit") options.PageHeight = global::OfficeIMO.OfficeLength.Points(180);
        PdfDocument document = PdfDocument.Create(options);
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++) body.Text("ScopedNote" + index.ToString("00"), style: Style());
        });
        void Before(PdfContentBuilder body) => body.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("FirstReference").WithFootnoteReference(note)
        }), style: Style());

        if (boundary == "PageExit") {
            // A distinct height identifies the page options that must remain in
            // force until this page group's final note fragment is painted.
            document.Compose(pages => pages.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(260), global::OfficeIMO.OfficeLength.Points(140)).Content(Before)));
            document.Content.Text("FollowingBody", style: Style());
        } else if (boundary == "PanelSection") {
            document.Content.Panel(body => {
                Before(body);
                body.Section("After", after => after.Text("FollowingBody", style: Style()),
                    new PdfSectionOptions { StartOnNewPage = true, IncludeHeading = false });
            }, new PdfPanelStyle { PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
        } else {
            Before(document.Content);
            if (boundary == "Section") document.Content.Section("After", body => body.Text("FollowingBody", style: Style()),
                new PdfSectionOptions { StartOnNewPage = true, IncludeHeading = false });
            else document.Compose(pages => pages.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(260), global::OfficeIMO.OfficeLength.Points(180))
                .Content(body => body.Text("FollowingBody", style: Style()))));
        }

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("ScopedNote24", pdf.GetPage(3).Text);
        Assert.Contains("FollowingBody", pdf.GetPage(4).Text);
        Assert.All(pdf.GetPages().Take(3), page => {
            Assert.Equal(140D, page.Height);
            Assert.DoesNotContain("FollowingBody", page.Text);
        });
        Assert.Equal(boundary is "Section" or "PanelSection" ? 140D : 180D, pdf.GetPage(4).Height);
        for (int index = 0; index < 25; index++) {
            string token = "ScopedNote" + index.ToString("00");
            Assert.Single(pdf.GetPages(), page => page.Text.Contains(token));
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ExplicitPageBreakFollowsTheLastFragmentOfPreviouslyReferencedNotes(
        bool preserveEmptyPage, bool insidePanel) {
        PdfDocument document = PdfDocument.Create(Options());
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++)
                body.Text("BarrierNote" + index.ToString("00"), style: Style());
        });
        PdfFootnoteDefinition followingNote = document.Content.DefineFootnote(body =>
            body.Text("FollowingNote", style: Style()));

        void AddBody(PdfContentBuilder body) {
            body.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("FirstReference").WithFootnoteReference(note)
            }), style: Style());
            body.PageBreak(preserveEmptyPage);
            body.Paragraph(paragraph => paragraph.Runs(new[] {
                PdfTextRun.Normal("FollowingReference").WithFootnoteReference(followingNote)
            }), style: Style());
        }
        if (insidePanel)
            document.Content.Panel(AddBody, new PdfPanelStyle {
                PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            });
        else AddBody(document.Content);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("FirstReference", pdf.GetPage(1).Text);
        Assert.Contains("BarrierNote24", pdf.GetPage(3).Text);
        Assert.Contains("FollowingReference", pdf.GetPage(4).Text);
        Assert.Contains("FollowingNote", pdf.GetPage(4).Text);
        for (int index = 0; index < 25; index++) {
            string marker = "BarrierNote" + index.ToString("00");
            Assert.Single(pdf.GetPages(), page => page.Text.Contains(marker));
        }
        Assert.DoesNotContain("BarrierNote", pdf.GetPage(4).Text);
        Assert.All(pdf.GetPages().Take(3), page => {
            Assert.DoesNotContain("FollowingReference", page.Text);
            Assert.DoesNotContain("FollowingNote", page.Text);
        });
    }
}
