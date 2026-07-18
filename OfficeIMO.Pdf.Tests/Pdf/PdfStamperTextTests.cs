using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfStamperTests {
    [Fact]
    public void StampText_AddsTextToSelectedPageAndPreservesOriginalContent() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampText(source, "APPROVED", new PdfTextStampOptions {
            PageNumbers = new[] { 2 },
            X = 72,
            Y = 700,
            FontSize = 16,
            Color = PdfColor.Black
        });

        using var pdf = PdfPigDocument.Open(new MemoryStream(stamped));
        Assert.Equal(2, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(stamped);
        string firstPage = Normalize(read.Pages[0].ExtractText());
        string secondPage = Normalize(read.Pages[1].ExtractText());

        Assert.Contains("Firstpagebody", firstPage);
        Assert.DoesNotContain("APPROVED", firstPage);
        Assert.Contains("Secondpagebody", secondPage);
        Assert.Contains("APPROVED", secondPage);
    }

    [Fact]
    public void StampText_UsesInclusivePageRangeSelection() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampText(source, "RANGE-STAMP", new PdfTextStampOptions()
            .UsePageRange(2, 2));

        var read = PdfReadDocument.Open(stamped);
        Assert.DoesNotContain("RANGE-STAMP", Normalize(read.Pages[0].ExtractText()));
        Assert.Contains("RANGE-STAMP", Normalize(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void StampText_UsesInclusivePageRangeListSelectionAndDeduplicatesOverlap() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.StampText(source, "RANGE-LIST-STAMP", new PdfTextStampOptions()
            .UsePageRanges(PdfPageRange.ParseMany("1-2,2")));

        var read = PdfReadDocument.Open(stamped);
        string firstPage = Normalize(read.Pages[0].ExtractText());
        string secondPage = Normalize(read.Pages[1].ExtractText());

        Assert.Contains("RANGE-LIST-STAMP", firstPage);
        Assert.Contains("RANGE-LIST-STAMP", secondPage);
        Assert.Equal(1, CountOccurrences(firstPage, "RANGE-LIST-STAMP"));
        Assert.Equal(1, CountOccurrences(secondPage, "RANGE-LIST-STAMP"));
    }

    [Fact]
    public void StampText_FlattensReferencedContentArraysBeforeAddingStampStream() {
        byte[] stamped = PdfStamper.StampText(BuildIndirectContentsArrayPdf(), "STAMP", new PdfTextStampOptions {
            X = 20,
            Y = 20,
            FontSize = 10
        });

        var document = PdfReadDocument.Open(stamped);
        var (objects, _) = PdfSyntax.ParseObjects(stamped);
        int pageObjectNumber = document.Pages[0].ObjectNumber;
        var page = Assert.IsType<PdfDictionary>(objects[pageObjectNumber].Value);
        var contents = Assert.IsType<PdfArray>(page.Items["Contents"]);

        Assert.Equal(3, contents.Items.Count);
        foreach (var item in contents.Items) {
            var reference = Assert.IsType<PdfReference>(item);
            Assert.IsType<PdfStream>(objects[reference.ObjectNumber].Value);
        }
    }

    [Fact]
    public void WatermarkText_AddsDefaultWatermarkToEveryPage() {
        byte[] source = BuildTwoPagePdf();

        byte[] stamped = PdfStamper.WatermarkText(source, "DRAFT");

        using var pdf = PdfPigDocument.Open(new MemoryStream(stamped));
        Assert.Equal(2, pdf.NumberOfPages);

        var read = PdfReadDocument.Open(stamped);
        Assert.Contains("DRAFT", Normalize(read.Pages[0].ExtractText()));
        Assert.Contains("DRAFT", Normalize(read.Pages[1].ExtractText()));
    }

    [Fact]
    public void WatermarkText_CentersUsingStandardFontGlyphWidths() {
        byte[] stamped = PdfStamper.WatermarkText(BuildTwoPagePdf(), "WWWW", new PdfTextStampOptions {
            Font = PdfStandardFont.TimesRoman,
            FontSize = 10,
            RotationDegrees = 0
        });

        string stampContent = FindContentStreamContaining(stamped, "<57575757> Tj");

        Assert.Matches(@"1 0 -?0 1 278\.\d+ 421 Tm", stampContent);
    }
}
