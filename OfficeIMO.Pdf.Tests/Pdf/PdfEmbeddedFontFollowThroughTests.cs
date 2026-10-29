using System;
using System.Collections.Generic;
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
            content.TextField("asset", width: 200, value: "NW-42");
            content.TextField("inspector", width: 200);
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
    public void StampTextEmbedsTheDefaultBodyFontWhenRequested() {
        if (!DefaultBodyFontInstalled()) return;
        PdfDocument source = PdfDocument.Create(pdf => pdf.Content(content => content.Paragraph(p => p.Text("Body"))));

        byte[] embedded = source.Stamp.Text("Page 1 of 1", new PdfTextStampOptions { X = 72, Y = 36, EmbedFont = true, Font = PdfStandardFont.Helvetica }).ToBytes();
        byte[] standard = source.Stamp.Text("Page 1 of 1", new PdfTextStampOptions { X = 72, Y = 36, EmbedFont = false, Font = PdfStandardFont.Helvetica }).ToBytes();

        Assert.Contains("/FontFile", PdfTestContentText.ReadDecoded(embedded), StringComparison.Ordinal);
        Assert.Contains("Page 1 of 1", PdfReadDocument.Open(embedded).ExtractText(), StringComparison.Ordinal);
        Assert.Matches(new Regex(@"/BaseFont\s*/Helvetica\b"), PdfTestContentText.ReadDecoded(standard));
    }

    [Fact]
    public void ListStyleWithoutMarkerColorKeepsTheDefaultListMarkerColor() {
        var accent = PdfColor.FromRgb(12, 34, 210);
        byte[] bytes = PdfDocument.Create()
            .DefaultListStyle(new PdfListStyle { MarkerColor = accent })
            .Bullets(new[] { "First", "Second" }, style: new PdfListStyle { ItemSpacing = 2 })
            .ToBytes();

        Assert.Matches(new Regex(@"0\.047\d* 0\.133\d* 0\.82\d* rg"), PdfTestContentText.ReadDecoded(bytes));
    }
}
