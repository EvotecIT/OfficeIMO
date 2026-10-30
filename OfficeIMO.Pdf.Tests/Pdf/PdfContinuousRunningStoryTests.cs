using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfContinuousRunningStoryTests {
    [Fact]
    public void ContinuousHeaderKeepsTheCurrentPhysicalPageWithoutForcingAnotherPage() {
        var document = Create();
        Line(document, "FIRST");
        document.BeginRunningContentSection(1, Presentation("HEADER_B"));
        Line(document, "SECOND");

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("HEADER_A", pdf.GetPage(1).Text);
        Assert.DoesNotContain("HEADER_B", pdf.GetPage(1).Text);
        Assert.Contains("FIRST", pdf.GetPage(1).Text);
        Assert.Contains("SECOND", pdf.GetPage(1).Text);
    }

    [Fact]
    public void NextPhysicalPageUsesASnapshotOfTheReplacementHeaderAcrossReplays() {
        var document = Create();
        Line(document, "FIRST");
        var next = Presentation("HEADER_B");
        document.BeginRunningContentSection(1, next);
        next.HeaderFormat = "MUTATED_AFTER_CAPTURE";
        Line(document, "SECOND");
        document.PageBreak();
        Line(document, "THIRD");

        for (int replay = 0; replay < 2; replay++) {
            using var pdf = PdfPigDocument.Open(document.ToBytes());
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("HEADER_A", pdf.GetPage(1).Text);
            Assert.DoesNotContain("HEADER_B", pdf.GetPage(1).Text);
            Assert.Contains("HEADER_B", pdf.GetPage(2).Text);
            Assert.DoesNotContain("HEADER_A", pdf.GetPage(2).Text);
            Assert.Contains("SECOND", pdf.GetPage(1).Text);
            Assert.Contains("THIRD", pdf.GetPage(2).Text);
            Assert.DoesNotContain("MUTATED_AFTER_CAPTURE", string.Concat(pdf.GetPages().Select(page => page.Text)));
        }
    }

    [Fact]
    public void NaturalParagraphOverflowUsesTheReplacementHeaderAndPreservesEveryBodyLine() {
        var document = Create();
        Line(document, "FIRST");
        document.BeginRunningContentSection(1, Presentation("HEADER_B"));
        document.Paragraph(paragraph => {
            for (int index = 0; index < 12; index++) {
                if (index != 0) paragraph.LineBreak();
                paragraph.Text("BODY" + index.ToString("D2"));
            }
        }, style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("HEADER_A", pdf.GetPage(1).Text);
        Assert.Contains("HEADER_B", pdf.GetPage(2).Text);
        Assert.DoesNotContain("HEADER_B", pdf.GetPage(1).Text);
        Assert.DoesNotContain("HEADER_A", pdf.GetPage(2).Text);
        string all = string.Concat(pdf.GetPages().Select(page => page.Text));
        for (int index = 0; index < 12; index++) {
            string marker = "BODY" + index.ToString("D2");
            Assert.Equal(1, all.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        }
    }

    [Fact]
    public void AKeptParagraphIsRemeasuredInTheReplacementRunningFrame() {
        var document = Create();
        document.Header(header => header.Content(content => content.Spacer(global::OfficeIMO.OfficeLength.Points(220)), 18));
        Line(document, "FIRST");
        var next = PdfDocument.Create(Presentation("HEADER_B"));
        next.Header(header => header.Content(content => content.Spacer(global::OfficeIMO.OfficeLength.Points(20)), 18));
        document.BeginRunningContentSection(1, next.Options);
        document.Paragraph(paragraph => paragraph.Text("NEXT_FRAME_BODY"), style: new PdfParagraphStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(12), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0),
            LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(80)), KeepTogether = true, WidowControl = false
        });

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FIRST", pdf.GetPage(1).Text);
        Assert.DoesNotContain("NEXT_FRAME_BODY", pdf.GetPage(1).Text);
        Assert.Contains("NEXT_FRAME_BODY", pdf.GetPage(2).Text);
    }

    [Fact]
    public void SuccessiveContinuousSectionsSelectTheLastHeaderAtTheNextPhysicalPage() {
        var document = Create();
        Line(document, "FIRST");
        document.BeginRunningContentSection(1, Presentation("HEADER_B"));
        Line(document, "SECOND");
        document.BeginRunningContentSection(2, Presentation("HEADER_C"));
        Line(document, "THIRD");
        document.PageBreak();
        Line(document, "FOURTH");

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("HEADER_A", pdf.GetPage(1).Text);
        Assert.Contains("HEADER_C", pdf.GetPage(2).Text);
        Assert.Contains("SECOND", pdf.GetPage(1).Text);
        Assert.Contains("THIRD", pdf.GetPage(1).Text);
        Assert.Contains("FOURTH", pdf.GetPage(2).Text);
        Assert.DoesNotContain("HEADER_B", string.Concat(pdf.GetPages().Select(page => page.Text)));
    }

    [Fact]
    public void AFinalLogicalMarkerDoesNotCreateARunningContentOnlyPage() {
        var document = Create();
        Line(document, "ONLY_BODY");
        document.BeginRunningContentSection(1, Presentation("HEADER_B"));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("HEADER_A", pdf.GetPage(1).Text);
        Assert.Contains("ONLY_BODY", pdf.GetPage(1).Text);
        Assert.DoesNotContain("HEADER_B", pdf.GetPage(1).Text);
    }

    [Fact]
    public void ExplicitPageSectionOwnsItsHeaderInsteadOfAnUnusedContinuousPresentation() {
        var document = Create();
        Line(document, "FIRST");
        document.BeginRunningContentSection(1, Presentation("HEADER_B"));
        document.Page(page => page.Header(header => header.Text("HEADER_C"))
            .Content(content => content.Paragraph(paragraph => paragraph.Text("EXPLICIT_BODY"))));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("HEADER_A", pdf.GetPage(1).Text);
        Assert.Contains("HEADER_C", pdf.GetPage(2).Text);
        Assert.Contains("EXPLICIT_BODY", pdf.GetPage(2).Text);
        Assert.DoesNotContain("HEADER_B", string.Concat(pdf.GetPages().Select(page => page.Text)));
    }

    [Fact]
    public void ReplacementPresentationCannotRelaxTheDocumentGeneratedPageLimit() {
        var options = Presentation("HEADER_A");
        options.MaxGeneratedPages = 1;
        var document = PdfDocument.Create(options);
        Line(document, "FIRST");
        var next = Presentation("HEADER_B");
        next.MaxGeneratedPages = 10;
        document.BeginRunningContentSection(1, next);
        document.PageBreak();
        Line(document, "SECOND");

        Assert.Throws<System.IO.InvalidDataException>(() => document.ToBytes());
    }

    private static PdfDocument Create() {
        var document = PdfDocument.Create(Presentation("HEADER_A"));
        document.BeginRunningContentSection(0);
        return document;
    }

    private static PdfCreateOptions Presentation(string header) => new() {
        PageWidth = global::OfficeIMO.OfficeLength.Points(240), PageHeight = global::OfficeIMO.OfficeLength.Points(320),
        MarginTop = global::OfficeIMO.OfficeLength.Points(32), MarginBottom = global::OfficeIMO.OfficeLength.Points(32), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
        ShowHeader = true, HeaderFormat = header, HeaderOffsetY = global::OfficeIMO.OfficeLength.Points(18)
    };

    private static PdfParagraphStyle Style() => new() {
        FontSize = global::OfficeIMO.OfficeLength.Points(10), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0),
        LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(24)), WidowControl = false
    };

    private static void Line(PdfDocument document, string value) =>
        document.Paragraph(paragraph => paragraph.Text(value), style: Style());
}
