using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfLogicalDocumentTests {
    [Fact]
    public void Load_BuildsLogicalPagesWithTextTablesAndImages() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical sample", author: "OfficeIMO")
            .Paragraph(p => p.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(CreateMinimalRgbPng(), 18, 18)
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.Equal("Logical sample", logical.Metadata.Title);
        Assert.Contains(page.TextBlocks, block => Normalize(block.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(page.TextBlocks, block =>
            block.Kind == PdfLogicalElementKind.ListItem &&
            Normalize(block.Text).Contains("Detectedlogicalbullet", StringComparison.Ordinal));
        Assert.Contains(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("A-100", StringComparison.Ordinal));

        PdfLogicalTable table = Assert.Single(page.Tables, item => item.Rows.Count >= 3 && item.Columns.Count >= 3);
        Assert.Contains(table.Rows, row => row.Count >= 3 &&
            Normalize(row[0]) == "A-100" &&
            Normalize(row[1]) == "Alpha" &&
            Normalize(row[2]) == "2");

        PdfLogicalImage image = Assert.Single(page.Images);
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("image/png", image.MimeType);

        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.Table);
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.Image);
    }

    [Fact]
    public void Load_GroupsWrappedTextLinesIntoParagraphs() {
        byte[] pdf = PdfDoc.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("This logical paragraph should wrap across multiple nearby PDF text lines so wrappers can start from paragraph-like objects."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "P-100", "Paragraph table text", "2" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 50, 100, 30 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfLogicalPage page = Assert.Single(PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages);

        PdfLogicalParagraph paragraph = Assert.Single(page.Paragraphs, item => item.Text.Contains("logical paragraph", StringComparison.Ordinal));
        Assert.True(paragraph.Lines.Count > 1);
        Assert.Contains("logical paragraph", paragraph.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("P-100", paragraph.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Load_ExposesSimpleAcroFormFields() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildHierarchicalFormPdf());

        Assert.True(logical.HasFormFields);
        Assert.Equal(new[] { "Person.Name", "AcceptTerms" }, logical.FormFields.Select(field => field.Name).ToArray());
        Assert.Equal("OfficeIMO", logical.FormFields[0].Value);
        Assert.Equal("Yes", logical.FormFields[1].Value);
    }

    [Fact]
    public void Load_ReadsStreamFromCurrentPosition() {
        byte[] source = PdfDoc.Create()
            .Paragraph(p => p.Text("Logical stream marker."))
            .ToBytes();
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        using var stream = new MemoryStream(prefix.Concat(source).ToArray());
        stream.Position = prefix.Length;

        PdfLogicalDocument logical = PdfLogicalDocument.Load(stream);

        Assert.Single(logical.Pages);
        Assert.Contains(logical.Pages[0].TextBlocks, block => block.Text.Contains("Logical stream marker", StringComparison.Ordinal));
    }

    private static string Normalize(string text) {
        return new string(text.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
    }

    private static byte[] BuildHierarchicalFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [6 0 R 8 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /T (Person) /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Tx /T (Name) /TU (Display name) /TM (ExportName) /V (OfficeIMO) /Ff 1 >>",
            "endobj",
            "8 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }
}
