using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfReadLayoutSmokeTests {
    [Fact]
    public void PdfReadDocument_ColumnAndStructuredApis_DoNotThrow() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        try {
            var pdf = PdfDoc.Create()
                .Meta(title: "Smoke")
                .H1("Header")
                .Paragraph(p => p.Text("Line one for extraction."))
                .Paragraph(p => p.Text("Line two for extraction."));

            pdf.Save(path);

            var doc = PdfReadDocument.Load(path);
            Assert.NotNull(doc);
            Assert.NotEmpty(doc.Pages);

            var text = doc.ExtractTextWithColumns();
            Assert.NotNull(text);

            var structured = doc.ExtractStructured();
            Assert.NotNull(structured.Lines);
            Assert.NotNull(structured.Toc);
            Assert.NotNull(structured.Lists);
            Assert.NotNull(structured.LeaderRows);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void PdfReadPage_ExtensionApis_DoNotThrow() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        try {
            var pdf = PdfDoc.Create()
                .Paragraph(p => p.Text("Page extension api smoke."));
            pdf.Save(path);

            var doc = PdfReadDocument.Load(path);
            Assert.NotEmpty(doc.Pages);

            var page = doc.Pages[0];
            var text = page.ExtractTextWithColumns(new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });
            Assert.NotNull(text);

            var structured = page.ExtractStructured(new PdfTextLayoutOptions());
            Assert.NotNull(structured);
            Assert.NotNull(structured.Lines);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }
}
