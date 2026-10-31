using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData(PdfFootnotePlacement.BottomOfPage, 7D)]
    [InlineData(PdfFootnotePlacement.BottomOfPage, 17D)]
    [InlineData(PdfFootnotePlacement.BeneathText, 7D)]
    [InlineData(PdfFootnotePlacement.BeneathText, 17D)]
    public void CompleteNoteCapturesFinalReservationAndMeasuresItsEntireNoticeFlow(
        PdfFootnotePlacement placement, double noticeSpacer) {
        PdfDocument Render(bool reserve, int noticeKind = 1) {
            var document = PdfDocument.Create(Options());
            Action<PdfContentBuilder>? notice = noticeKind switch {
                1 => body => body.Text("UnpaintedNotice", style: Style()).Spacer(global::OfficeIMO.OfficeLength.Points(noticeSpacer)),
                2 => _ => { },
                _ => null
            };
            var options = new PdfFootnoteOptions {
                Placement = placement, ReserveContinuationNoticeSpaceOnFinalPage = reserve,
                ContinuationNotice = notice
            };
            var note = document.Content.DefineFootnote(body => {
                // Both true and false must be captured before this callback runs.
                options.ReserveContinuationNoticeSpaceOnFinalPage = !reserve;
                body.Text("CompleteShortNote", style: Style());
            }, options);
            document.Content.Paragraph(p => p.Runs(new[] {
                PdfTextRun.Normal("ShortReference").WithFootnoteReference(note)
            }), style: Style());
            document.Content.Spacer(global::OfficeIMO.OfficeLength.Points(65D)).Text("FollowingBody", style: Style());
            return document;
        }

        using var control = PdfPigDocument.Open(Render(reserve: false).ToBytes());
        Assert.Equal(1, control.NumberOfPages);
        Assert.Contains("FollowingBody", control.GetPage(1).Text, StringComparison.Ordinal);
        var reserved = Render(reserve: true);
        foreach (byte[] bytes in new[] { reserved.ToBytes(), reserved.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("ShortReference", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("CompleteShortNote", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.DoesNotContain("FollowingBody", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("FollowingBody", pdf.GetPage(2).Text, StringComparison.Ordinal);
            Assert.All(pdf.GetPages(), page => Assert.DoesNotContain("UnpaintedNotice", page.Text, StringComparison.Ordinal));
            if (placement == PdfFootnotePlacement.BottomOfPage) {
                double shift = TableMarkerBaseline(pdf, 1, "CompleteShortNote") -
                    TableMarkerBaseline(control, 1, "CompleteShortNote");
                Assert.Equal(10D + noticeSpacer, shift, 3);
            }
        }

        // An absent or explicitly empty selected notice contributes zero area,
        // even when the final-page reservation policy was captured as true.
        foreach (int noticeKind in new[] { 0, 2 }) {
            using var noNotice = PdfPigDocument.Open(Render(reserve: true, noticeKind: noticeKind).ToBytes());
            Assert.Equal(1, noNotice.NumberOfPages);
            Assert.Contains("FollowingBody", noNotice.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Equal(TableMarkerBaseline(control, 1, "CompleteShortNote"),
                TableMarkerBaseline(noNotice, 1, "CompleteShortNote"), 3);
        }
    }

    [Theory]
    [InlineData(PdfFootnotePlacement.BottomOfPage)]
    [InlineData(PdfFootnotePlacement.BeneathText)]
    public void ContinuedNoteReservesMeasuredNoticeSpaceOnItsFinalPageWithoutPaintingIt(PdfFootnotePlacement placement) {
        PdfDocument Render(bool reserve) {
            var document = PdfDocument.Create(Options());
            var note = document.Content.DefineFootnote(body => {
                for (int index = 0; index < 24; index++) body.Text("ReservedLong" + index.ToString("00"), style: Style());
            }, new PdfFootnoteOptions {
                Placement = placement, ReserveContinuationNoticeSpaceOnFinalPage = reserve,
                ContinuationNotice = body => body.Text("MeasuredNotice", style: Style()).Spacer(global::OfficeIMO.OfficeLength.Points(7D))
            });
            document.Content.Paragraph(p => p.Runs(new[] {
                PdfTextRun.Normal("LongReference").WithFootnoteReference(note)
            }), style: Style());
            document.Content.Text("FollowingLongBody", style: Style());
            return document;
        }

        using var control = PdfPigDocument.Open(Render(reserve: false).ToBytes());
        Assert.Equal(3, control.NumberOfPages);
        Assert.Contains("ReservedLong23", control.GetPage(3).Text, StringComparison.Ordinal);
        Assert.Contains("FollowingLongBody", control.GetPage(3).Text, StringComparison.Ordinal);
        var document = Render(reserve: true);
        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(4, pdf.NumberOfPages);
            Assert.Contains("ReservedLong23", pdf.GetPage(4).Text, StringComparison.Ordinal);
            Assert.Contains("FollowingLongBody", pdf.GetPage(4).Text, StringComparison.Ordinal);
            foreach (var page in pdf.GetPages().Take(3))
                Assert.Single(page.GetWords(), word => word.Text == "MeasuredNotice");
            Assert.DoesNotContain("MeasuredNotice", pdf.GetPage(4).Text, StringComparison.Ordinal);
            for (int index = 0; index < 24; index++) {
                string token = "ReservedLong" + index.ToString("00");
                Assert.Single(pdf.GetPages(), page => page.GetWords().Any(word => word.Text == token));
            }
            if (placement == PdfFootnotePlacement.BottomOfPage)
                Assert.Equal(17D, TableMarkerBaseline(pdf, 4, "ReservedLong23") -
                    TableMarkerBaseline(control, 3, "ReservedLong23"), 3);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FirstAdmittedTableNoteOwnsFinalReservationWithoutChargingItAgain(bool firstAdmittedReserves) {
        var options = Options();
        options.PageHeight = global::OfficeIMO.OfficeLength.Points(260D);
        var document = PdfDocument.Create(options);
        // The right-cell note is admitted first, although the spanned left-cell
        // reference precedes it in authored source order and is admitted later.
        var firstAdmitted = document.Content.DefineFootnote(body => body.Text("FirstAdmittedNote", style: Style()),
            new PdfFootnoteOptions {
                ReserveContinuationNoticeSpaceOnFinalPage = firstAdmittedReserves,
                ContinuationNotice = body => body.Text("OwnerNotice", style: Style()).Spacer(global::OfficeIMO.OfficeLength.Points(7D))
            });
        var earlierAuthored = document.Content.DefineFootnote(body => body.Text("EarlierAuthoredNote", style: Style()),
            new PdfFootnoteOptions {
                ReserveContinuationNoticeSpaceOnFinalPage = !firstAdmittedReserves,
                ContinuationNotice = body => body.Text("UnselectedNotice", style: Style()).Spacer(global::OfficeIMO.OfficeLength.Points(25D))
            });
        var style = FootnoteTableStyle();
        style.MinimumBodyRowsOnFirstPage = 0;
        style.MinimumBodyRowsOnLastPage = 0;
        document.Content.Table(new[] {
            new[] {
                new PdfTableCell(TableReferenceLines(12, 8, "LRef", earlierAuthored), rowSpan: 2),
                new PdfTableCell(new[] { PdfTextRun.Normal("RRef").WithFootnoteReference(firstAdmitted), PdfTextRun.Normal("\nRightFirstTail") })
            },
            new[] { new PdfTableCell(string.Join("\n", Enumerable.Range(1, 10).Select(index => "RightTail" + index))) }
        }, style: style);

        var baselineDocument = PdfDocument.Create(options);
        var baselineNote = baselineDocument.Content.DefineFootnote(body => body.Text("BaselineNote", style: Style()));
        baselineDocument.Content.Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("BaselineReference").WithFootnoteReference(baselineNote)
        }), style: Style());
        using var baseline = PdfPigDocument.Open(baselineDocument.ToBytes());
        double ordinaryBottomNote = TableMarkerBaseline(baseline, 1, "BaselineNote");

        foreach (byte[] bytes in new[] { document.ToBytes(), document.ToBytes() }) {
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(1, pdf.NumberOfPages);
            Assert.Contains("LRef", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("RRef", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "EarlierAuthoredNote");
            Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "FirstAdmittedNote");
            Assert.DoesNotContain("OwnerNotice", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.DoesNotContain("UnselectedNotice", pdf.GetPage(1).Text, StringComparison.Ordinal);
            double earlier = TableMarkerBaseline(pdf, 1, "EarlierAuthoredNote");
            double later = TableMarkerBaseline(pdf, 1, "FirstAdmittedNote");
            Assert.Equal(10D, earlier - later, 3);
            Assert.Equal(firstAdmittedReserves ? 17D : 0D, later - ordinaryBottomNote, 3);
        }
    }

    [Theory]
    [InlineData(PdfFootnotePlacement.BottomOfPage)]
    [InlineData(PdfFootnotePlacement.BeneathText)]
    public void ImpossibleFinalReservationFailsBeforeProducingNoticeOnlyPages(PdfFootnotePlacement placement) {
        var options = Options();
        options.MaxGeneratedPages = 2;
        var document = PdfDocument.Create(options);
        var note = document.Content.DefineFootnote(body => body.Text("OtherwiseCompleteNote", style: Style()),
            new PdfFootnoteOptions {
                Placement = placement, ReserveContinuationNoticeSpaceOnFinalPage = true,
                ContinuationNotice = body => body.Spacer(global::OfficeIMO.OfficeLength.Points(100D))
            });
        document.Content.Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("ImpossibleReference").WithFootnoteReference(note)
        }), style: Style());
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }
}
