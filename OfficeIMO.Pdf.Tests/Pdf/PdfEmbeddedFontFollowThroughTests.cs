using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Filled form values and stamps keep the embedded body font of the document they change.</summary>
public class PdfEmbeddedFontFollowThroughTests {
    private static bool DefaultBodyFontInstalled() =>
        PdfCreateOptions.TryResolveDefaultEmbeddedFontFamily(PdfStandardFont.Helvetica, out PdfEmbeddedFontFamily? family) && family != null;

    [Fact]
    public void FilledValuesReuseTheInstalledFontBehindTheEmbeddedFormFont() {
        if (!DefaultBodyFontInstalled()) return;
        PdfDocument template = PdfDocument.Create(pdf => pdf.Content(content => {
            content.Paragraph(p => p.Text("Inspector"));
            content.TextField("asset", width: OfficeLength.Points(200), value: "NW-42");
            content.TextField("inspector", width: OfficeLength.Points(200));
        }), new PdfCreateOptions { EmbedDefaultFonts = true });

        PdfDocument completed = template.Forms.Fill(new Dictionary<string, string> { ["inspector"] = "Marta Nowak" });
        byte[] completedBytes = completed.ToBytes();
        byte[] recordBytes = completed.Forms.Flatten().ToBytes();

        foreach (byte[] bytes in new[] { completedBytes, recordBytes }) {
            Assert.DoesNotMatch(new Regex(@"/BaseFont\s*/Helvetica\b"), PdfTestContentText.ReadDecoded(bytes));
        }
        Assert.Contains("Marta Nowak", PdfReadDocument.Open(recordBytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void FieldsFilledTogetherShareOneSubsetTaggedFont() {
        if (!DefaultBodyFontInstalled()) return;
        PdfDocument template = PdfDocument.Create(pdf => pdf.Content(content => {
            content.TextField("first", width: OfficeLength.Points(200));
            content.TextField("second", width: OfficeLength.Points(200));
            content.TextField("third", width: OfficeLength.Points(200));
        }), new PdfCreateOptions { EmbedDefaultFonts = true });

        byte[] bytes = template.Forms.Fill(new Dictionary<string, string> {
            ["first"] = "Marta Nowak", ["second"] = "Quality review", ["third"] = "Zebra crossing" }).ToBytes();
        string raw = PdfTestContentText.ReadDecoded(bytes);

        MatchCollection fontFiles = Regex.Matches(raw, @"/FontFile2");
        Assert.True(fontFiles.Count <= 3, "Expected the filled values to share one embedded appearance font.");
        Assert.DoesNotMatch(new Regex(@"/BaseFont\s*/Aptos(?![-\w+])"), raw);
    }

    [Theory]
    [InlineData(PdfAlign.Left)]
    [InlineData(PdfAlign.Center)]
    [InlineData(PdfAlign.Right)]
    public void StampAlignmentAnchorsTheTextAtX(PdfAlign align) {
        PdfDocument source = PdfDocument.Create(pdf => pdf.Content(content => content.Paragraph(p => p.Text("Body"))));
        byte[] bytes = source.Stamp.Text("Page 1 of 1", new PdfTextStampOptions { X = OfficeLength.Points(300), Y = OfficeLength.Points(36), Align = align, EmbedFont = false }).ToBytes();

        PdfTextSpan span = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans().First(item => item.Text.Contains("Page"));
        double expected = align == PdfAlign.Left ? span.X : align == PdfAlign.Center ? span.X + span.Advance / 2 : span.X + span.Advance;
        Assert.InRange(expected, 299, 301);
    }

    [Fact]
    public void StampTextEmbedsTheDefaultBodyFontWhenRequested() {
        if (!DefaultBodyFontInstalled()) return;
        PdfDocument source = PdfDocument.Create(pdf => pdf.Content(content => content.Paragraph(p => p.Text("Body"))));

        byte[] embedded = source.Stamp.Text("Page 1 of 1", new PdfTextStampOptions { X = OfficeLength.Points(72), Y = OfficeLength.Points(36), EmbedFont = true, Font = PdfStandardFont.Helvetica }).ToBytes();
        byte[] standard = source.Stamp.Text("Page 1 of 1", new PdfTextStampOptions { X = OfficeLength.Points(72), Y = OfficeLength.Points(36), EmbedFont = false, Font = PdfStandardFont.Helvetica }).ToBytes();

        Assert.Contains("/FontFile", PdfTestContentText.ReadDecoded(embedded), StringComparison.Ordinal);
        Assert.Contains("Page 1 of 1", PdfReadDocument.Open(embedded).ExtractText(), StringComparison.Ordinal);
        Assert.Matches(new Regex(@"/BaseFont\s*/Helvetica\b"), PdfTestContentText.ReadDecoded(standard));
    }

    [Fact]
    public void ListStyleWithoutMarkerColorKeepsTheDefaultListMarkerColor() {
        var accent = OfficeColor.FromRgb(12, 34, 210);
        byte[] bytes = PdfDocument.Create()
            .DefaultListStyle(new PdfListStyle { MarkerColor = accent })
            .Bullets(new[] { "First", "Second" }, style: new PdfListStyle { ItemSpacing = OfficeLength.Points(2) })
            .ToBytes();

        Assert.Matches(new Regex(@"0\.047\d* 0\.133\d* 0\.82\d* rg"), PdfTestContentText.ReadDecoded(bytes));
    }
}
