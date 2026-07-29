using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfImportTests {
    [Fact]
    public void PdfToRtf_PreservesLogicalRunColorSizeAndFontStyle() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph
                .Color(PdfCore.PdfColor.FromRgb(255, 0, 0))
                .FontSize(14)
                .Bold("Bold red")
                .Color(PdfCore.PdfColor.FromRgb(0, 0, 255))
                .FontSize(9)
                .Italic(" italic blue"))
            .ToBytes();

        RtfDocument document = LoadSemanticPdf(pdf).ToRtfDocument();
        RtfRun red = document.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .First(run => run.Text.Contains("Bold red", StringComparison.Ordinal));
        RtfRun blue = document.Paragraphs
            .SelectMany(paragraph => paragraph.Runs)
            .First(run => run.Text.Contains("italic blue", StringComparison.Ordinal));

        Assert.True(red.Bold);
        Assert.Equal(14D, red.FontSize);
        Assert.NotNull(red.ForegroundColorIndex);
        RtfColor redColor = document.Colors[red.ForegroundColorIndex!.Value - 1];
        Assert.Equal((byte)255, redColor.Red);
        Assert.Equal((byte)0, redColor.Green);
        Assert.Equal((byte)0, redColor.Blue);

        Assert.True(blue.Italic);
        Assert.Equal(9D, blue.FontSize);
        Assert.NotNull(blue.ForegroundColorIndex);
        RtfColor blueColor = document.Colors[blue.ForegroundColorIndex!.Value - 1];
        Assert.Equal((byte)0, blueColor.Red);
        Assert.Equal((byte)0, blueColor.Green);
        Assert.Equal((byte)255, blueColor.Blue);
    }

    [Fact]
    public void PdfToRtf_PreservesStyledListRunsWithoutRepeatingMarker() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph
                .Text("\u2022 ")
                .Color(PdfCore.PdfColor.FromRgb(255, 0, 0))
                .Bold("Bold red")
                .Color(PdfCore.PdfColor.FromRgb(0, 0, 255))
                .Italic(" italic blue"))
            .ToBytes();
        PdfCore.PdfLogicalDocument logical = LoadSemanticPdf(pdf);
        PdfCore.PdfLogicalListItem logicalItem = Assert.Single(logical.ListItems);
        Assert.Equal(logicalItem.Text, string.Concat(logicalItem.Runs.Select(run => run.Text)));

        RtfDocument document = logical.ToRtfDocument();
        RtfParagraph listItem = Assert.Single(document.Paragraphs);
        Assert.Equal("Bold red italic blue", listItem.ToPlainText());
        Assert.Equal(RtfListKind.Bullet, listItem.ListKind);
        RtfRun red = listItem.Runs.First(run => run.Text.Contains("Bold red", StringComparison.Ordinal));
        RtfRun blue = listItem.Runs.First(run => run.Text.Contains("italic blue", StringComparison.Ordinal));
        Assert.True(red.Bold);
        Assert.True(blue.Italic);
        Assert.NotNull(red.ForegroundColorIndex);
        Assert.NotNull(blue.ForegroundColorIndex);
    }

    [Fact]
    public void PdfBytes_ToRtfDocument_Imports_Metadata_Headings_Lists_Paragraphs_And_PageBreaks() {
        byte[] pdf = CreateSemanticPdf();

        RtfDocument document = LoadSemanticPdf(pdf).ToRtfDocument(CreateImportOptions());

        Assert.Equal("PDF Import Title", document.Info.Title);
        Assert.Equal("OfficeIMO", document.Info.Author);
        Assert.Equal("PDF to RTF", document.Info.Subject);
        Assert.Equal("pdf,rtf", document.Info.Keywords);
        Assert.Equal(7200, document.PageSetup.PaperWidthTwips);
        Assert.Equal(7200, document.PageSetup.PaperHeightTwips);
        Assert.Contains(document.Styles, style => style.Name == "Heading 1" && style.Bold == true);

        RtfParagraph heading = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        Assert.Equal(0, heading.OutlineLevel);
        Assert.True(heading.Runs[0].Bold);

        Assert.Contains(document.Paragraphs, paragraph => paragraph.ToPlainText().Contains("semantic paragraph", StringComparison.Ordinal));

        RtfParagraph bullet = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "First bullet");
        Assert.Equal(RtfListKind.Bullet, bullet.ListKind);
        Assert.Equal(0, bullet.ListLevel);
        Assert.NotNull(bullet.ListText);

        RtfParagraph numbered = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "First numbered");
        Assert.Equal(RtfListKind.Decimal, numbered.ListKind);
        Assert.NotNull(numbered.ListText);

        RtfParagraph secondPage = Assert.Single(document.Paragraphs, paragraph => paragraph.ToPlainText() == "Second page body.");
        Assert.True(secondPage.PageBreakBefore);
    }

    [Fact]
    public void Pdf_ToRtf_Serializes_And_RoundTrips_Imported_Text() {
        byte[] pdf = CreateSemanticPdf();

        string rtf = LoadSemanticPdf(pdf).ToRtfDocument(CreateImportOptions()).ToRtf();
        RtfDocument roundTrip = RtfDocument.Read(rtf).Document;

        Assert.Contains("Clinical Summary", rtf, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ToPlainText() == "Second page body.");
        Assert.Contains(roundTrip.Paragraphs, paragraph => paragraph.ListKind == RtfListKind.Bullet && paragraph.ToPlainText() == "First bullet");
    }

    [Fact]
    public void Pdf_Stream_File_And_Save_Apis_Import_Rtf() {
        byte[] pdf = CreateSemanticPdf();
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-rtf-import-" + Guid.NewGuid().ToString("N"));
        string pdfPath = Path.Combine(directory, "source.pdf");
        string rtfPath = Path.Combine(directory, "output.rtf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(pdfPath, pdf);

            using MemoryStream pdfStream = new MemoryStream(pdf);
            RtfDocument fromStream = PdfCore.PdfLogicalDocument
                .Load(pdfStream, CreateLayoutOptions())
                .ToRtfDocument(CreateImportOptions());
            Assert.Contains(fromStream.Paragraphs, paragraph => paragraph.ToPlainText() == "First bullet");

            RtfDocument fromFile = PdfCore.PdfLogicalDocument
                .Load(pdfPath, CreateLayoutOptions())
                .ToRtfDocument(CreateImportOptions());
            Assert.Contains(fromFile.Paragraphs, paragraph => paragraph.ToPlainText() == "Second page body.");

            fromFile.Save(rtfPath, encoding: Encoding.UTF8);
            RtfDocument saved = RtfDocument.Load(rtfPath, encoding: Encoding.UTF8).Document;
            Assert.Contains(saved.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PdfRtfImportOptions_Clone_IsIndependent() {
        var options = new PdfRtfImportOptions {
            PreservePageBreaks = false,
            IncludeMetadata = false
        };

        PdfRtfImportOptions clone = options.Clone();
        clone.PreservePageBreaks = true;
        clone.IncludeMetadata = true;

        Assert.False(options.PreservePageBreaks);
        Assert.False(options.IncludeMetadata);
        Assert.True(clone.PreservePageBreaks);
        Assert.True(clone.IncludeMetadata);
    }

    [Fact]
    public async Task PdfDocument_RtfFacade_ReturnsDiagnosticsAndSupportsSyncAndAsyncSave() {
        byte[] pdf = CreateSemanticPdf();
        PdfCore.PdfDocument opened = PdfCore.PdfDocument.Open(pdf);

        PdfRtfConversionResult result = opened.ToRtfDocumentResult(CreateImportOptions());
        Assert.False(result.HasLoss);
        Assert.Contains(result.Value.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");

        using var sync = new MemoryStream();
        PdfRtfConversionReport syncReport = opened.SaveAsRtf(sync, CreateImportOptions());
        Assert.False(syncReport.HasLoss);
        Assert.NotEmpty(sync.ToArray());

        using var asyncOutput = new MemoryStream();
        PdfRtfConversionReport asyncReport = await opened.SaveAsRtfAsync(asyncOutput, CreateImportOptions());
        Assert.False(asyncReport.HasLoss);
        Assert.NotEmpty(asyncOutput.ToArray());
    }

    [Fact]
    public void PdfToRtf_ReportsDetectedTablesThatCannotBeReconstructed() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Due", "2026-06-30")
            })
            .ToBytes();
        PdfCore.PdfLogicalDocument logical = LoadSemanticPdf(pdf);
        Assert.NotEmpty(logical.Tables);

        PdfRtfConversionResult result = logical.ToRtfDocumentResult();

        Assert.True(result.HasLoss);
        PdfCore.PdfConversionWarning warning = Assert.Single(
            result.Report.Warnings,
            item => item.Code == "TABLES_NOT_IMPORTED");
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.Equal(logical.Tables.Count.ToString(), warning.Details["count"]);
    }

    private static PdfRtfImportOptions CreateImportOptions() => new PdfRtfImportOptions();

    private static PdfCore.PdfLogicalDocument LoadSemanticPdf(byte[] pdf) =>
        PdfCore.PdfLogicalDocument.Load(pdf, CreateLayoutOptions());

    private static PdfCore.PdfTextLayoutOptions CreateLayoutOptions() => new PdfCore.PdfTextLayoutOptions {
        ForceSingleColumn = true
    };

    private static byte[] CreateSemanticPdf() =>
        PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 360,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "PDF Import Title", author: "OfficeIMO", subject: "PDF to RTF", keywords: "pdf,rtf")
            .H1("Clinical Summary")
            .Paragraph(p => p.Text("This semantic paragraph should become one imported RTF paragraph."))
            .Bullets(new[] { "First bullet", "Second bullet" })
            .Numbered(new[] { "First numbered", "Second numbered" }, startNumber: 3)
            .PageBreak()
            .Paragraph(p => p.Text("Second page body."))
            .ToBytes();
}
