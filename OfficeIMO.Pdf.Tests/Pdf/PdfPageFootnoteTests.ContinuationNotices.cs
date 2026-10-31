using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData(PdfFootnotePlacement.BottomOfPage)]
    [InlineData(PdfFootnotePlacement.BeneathText)]
    public void ContinuationNoticeReservesSpaceAfterFragmentsAndIsAbsentFromFinalPage(PdfFootnotePlacement placement) {
        PdfDocument document = PdfDocument.Create(Options());
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++) body.Text("NoticeNote" + index.ToString("00"), style: Style());
        }, new PdfFootnoteOptions {
            Placement = placement,
            ContinuationNotice = body => body.Text("ContinuedNotice", style: Style())
        });
        document.Content.Paragraph(body => body.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        document.Content.Text("FollowingBody", style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("NoticeNote07", pdf.GetPage(1).Text);
        Assert.DoesNotContain("NoticeNote08", pdf.GetPage(1).Text);
        Assert.Contains("NoticeNote16", pdf.GetPage(2).Text);
        Assert.DoesNotContain("NoticeNote17", pdf.GetPage(2).Text);
        Assert.Contains("NoticeNote24", pdf.GetPage(3).Text);
        Assert.Contains("FollowingBody", pdf.GetPage(3).Text);
        foreach (var page in pdf.GetPages().Take(2)) {
            var notice = Assert.Single(page.GetWords(), word => word.Text == "ContinuedNotice");
            var lastNote = page.GetWords().Where(word => word.Text.StartsWith("NoticeNote", StringComparison.Ordinal))
                .OrderBy(word => word.BoundingBox.Bottom).First();
            Assert.True(notice.BoundingBox.Top <= lastNote.BoundingBox.Bottom + .001D);
            Assert.True(notice.BoundingBox.Bottom >= 19D);
            Assert.DoesNotContain("FollowingBody", page.Text);
        }
        Assert.DoesNotContain("ContinuedNotice", pdf.GetPage(3).Text);
        for (int index = 0; index < 25; index++) {
            string token = "NoticeNote" + index.ToString("00");
            Assert.Single(pdf.GetPages(), page => page.Text.Contains(token));
        }
    }

    [Fact]
    public void CompleteNoteDoesNotReserveOrPaintAnUnusedNotice() {
        PdfDocument document = PdfDocument.Create(Options());
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 9; index++) body.Text("CompleteNote" + index, style: Style());
        }, new PdfFootnoteOptions {
            ContinuationNotice = body => body.Text("UnusedNotice", style: new PdfParagraphStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(8), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(100)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
        });
        document.Content.Paragraph(body => body.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Single(pdf.GetPages());
        Assert.Contains("CompleteNote8", pdf.GetPage(1).Text);
        Assert.DoesNotContain("UnusedNotice", pdf.GetPage(1).Text);
    }

    [Fact]
    public void NoticeCallbacksAndRetainedBuildersAreCapturedOnceBeforeSerialization() {
        PdfDocument document = PdfDocument.Create(Options());
        int calls = 0;
        PdfContentBuilder? retained = null;
        var options = new PdfFootnoteOptions {
            ContinuationNotice = body => { calls++; retained = body; body.Text("CapturedNotice", style: Style()); }
        };
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            options.ContinuationNotice = changed => changed.Text("ChangedNotice", style: Style());
            for (int index = 0; index < 12; index++) body.Text("CapturedNote" + index, style: Style());
        }, options);
        Assert.NotNull(retained);
        // The retained builder is document-bound. Once capture ends, later calls
        // add ordinary body content rather than extending the frozen notice.
        retained!.Text("LateBody", style: Style());
        document.Content.Paragraph(body => body.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("CapturedNotice", pdf.GetPage(1).Text);
            Assert.DoesNotContain("CapturedNotice", pdf.GetPage(2).Text);
            Assert.Single(pdf.GetPages(), page => page.Text.Contains("LateBody"));
            Assert.True(pdf.GetPage(1).Text.IndexOf("LateBody", StringComparison.Ordinal) <
                pdf.GetPage(1).Text.IndexOf("Reference", StringComparison.Ordinal));
            Assert.All(pdf.GetPages(), page => {
                Assert.DoesNotContain("ChangedNotice", page.Text);
            });
        }
        Assert.Equal(1, calls);
    }

    [Fact]
    public void FirstAdmittedNoteSelectsOneNoticeForTheWholePage() {
        PdfDocument document = PdfDocument.Create(Options());
        PdfFootnoteDefinition first = document.Content.DefineFootnote(body => body.Text("ShortNote", style: Style()),
            new PdfFootnoteOptions { ContinuationNotice = body => body.Text("FirstNotice", style: Style()) });
        PdfFootnoteDefinition second = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++) body.Text("SecondNote" + index.ToString("00"), style: Style());
        }, new PdfFootnoteOptions { ContinuationNotice = body => body.Text("SecondNotice", style: Style()) });
        document.Content.Paragraph(body => body.Runs(new[] {
            PdfTextRun.Normal("First").WithFootnoteReference(first),
            PdfTextRun.Normal("Second").WithFootnoteReference(second)
        }), style: Style());
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "FirstNotice");
        Assert.DoesNotContain("SecondNotice", pdf.GetPage(1).Text);
        Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "SecondNotice");
        Assert.DoesNotContain("FirstNotice", pdf.GetPage(2).Text);
        Assert.DoesNotContain("Notice", pdf.GetPage(3).Text);
        Assert.Single(pdf.GetPages(), page => page.Text.Contains("ShortNote"));
        for (int index = 0; index < 25; index++) {
            string token = "SecondNote" + index.ToString("00");
            Assert.Single(pdf.GetPages(), page => page.Text.Contains(token));
        }
    }

    [Fact]
    public void ImpossibleContinuationNoticeFailsWithoutGeneratingUnboundedPages() {
        PdfCreateOptions options = Options();
        options.MaxGeneratedPages = 2;
        PdfDocument document = PdfDocument.Create(options);
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++) body.Text("LongNote" + index, style: Style());
        }, new PdfFootnoteOptions {
            ContinuationNotice = body => body.Spacer(global::OfficeIMO.OfficeLength.Points(100))
        });
        document.Content.Paragraph(body => body.Runs(new[] { PdfTextRun.Normal("Reference").WithFootnoteReference(note) }), style: Style());
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void ContinuationNoticeCannotIntroduceRecursiveReferences() {
        PdfDocument document = PdfDocument.Create(Options());
        PdfFootnoteDefinition nested = document.Content.DefineFootnote(body => body.Text("Nested", style: Style()));
        Assert.Throws<NotSupportedException>(() => document.Content.DefineFootnote(body => body.Text("Outer", style: Style()),
            new PdfFootnoteOptions {
                ContinuationNotice = body => body.Paragraph(paragraph => paragraph.Runs(new[] {
                    PdfTextRun.Normal("NestedReference").WithFootnoteReference(nested)
                }), style: Style())
            }));
    }
}
