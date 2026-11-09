using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPdfImportTests {
    [Fact]
    public void PdfToRtf_PreservesLogicalRunColorSizeAndFontStyle() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => {
                paragraph.AddText("Bold red").SetBold().SetColor(global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0)).SetFontSize(global::OfficeIMO.OfficeLength.Points(14));
                paragraph.AddText(" italic blue").SetItalic().SetColor(global::OfficeIMO.OfficeColor.FromRgb(0, 0, 255)).SetFontSize(global::OfficeIMO.OfficeLength.Points(9));
            })
            .ToBytes();

        RtfDocument document = LoadSemanticPdf(pdf).ToRtf();
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
            .Paragraph(paragraph => {
                paragraph.AddText("\u2022 ");
                paragraph.AddText("Bold red").SetBold().SetColor(global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0));
                paragraph.AddText(" italic blue").SetItalic().SetColor(global::OfficeIMO.OfficeColor.FromRgb(0, 0, 255));
            })
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = LoadSemanticPdf(pdf);
        PdfCore.PdfLogicalListItem logicalItem = Assert.Single(logical.ListItems);
        Assert.Equal(logicalItem.Text, string.Concat(logicalItem.Runs.Select(run => run.Text)));

        RtfDocument document = logical.ToRtf();
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

        RtfDocument document = LoadSemanticPdf(pdf).ToRtf(CreateImportOptions());

        Assert.Equal("PDF Import Title", document.Properties.Title);
        Assert.Equal("OfficeIMO", document.Properties.Author);
        Assert.Equal("PDF to RTF", document.Properties.Subject);
        Assert.Equal("pdf,rtf", document.Properties.Keywords);
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

        string rtf = LoadSemanticPdf(pdf).ToRtf(CreateImportOptions()).ToRtf();
        RtfDocument roundTrip = RtfDocument.ParseResult(rtf).Document;

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
            RtfDocument fromStream = PdfCore.PdfDocumentReadResult
                .Load(pdfStream, CreateLayoutOptions())
                .ToRtf(CreateImportOptions());
            Assert.Contains(fromStream.Paragraphs, paragraph => paragraph.ToPlainText() == "First bullet");

            RtfDocument fromFile = PdfCore.PdfDocumentReadResult
                .Load(pdfPath, CreateLayoutOptions())
                .ToRtf(CreateImportOptions());
            Assert.Contains(fromFile.Paragraphs, paragraph => paragraph.ToPlainText() == "Second page body.");

            fromFile.Save(rtfPath, encoding: Encoding.UTF8);
            RtfDocument saved = RtfDocument.LoadResult(rtfPath, encoding: Encoding.UTF8).Document;
            Assert.Contains(saved.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PdfToRtfOptions_Clone_IsIndependent() {
        var options = new PdfRtfExportOptions {
            PreservePageBreaks = false,
            IncludeMetadata = false
        };

        PdfRtfExportOptions clone = options.Clone();
        clone.PreservePageBreaks = true;
        clone.IncludeMetadata = true;

        Assert.False(options.PreservePageBreaks);
        Assert.False(options.IncludeMetadata);
        Assert.True(clone.PreservePageBreaks);
        Assert.True(clone.IncludeMetadata);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PdfDocument_RtfFacade_ReturnsDiagnosticsAndSupportsSyncAndAsyncSave(bool logicalSource) {
        byte[] pdf = CreateSemanticPdf();
        using PdfCore.PdfDocument opened = PdfCore.PdfDocument.Load(pdf);
        PdfCore.PdfDocumentReadResult logical = LoadSemanticPdf(pdf);

        OfficeConversionResult<RtfDocument> result = logicalSource
            ? await logical.ToRtfAsync(CreateImportOptions())
            : await opened.ToRtfAsync(CreateImportOptions());
        using RtfDocument converted = result.Value;
        PdfRtfConversionReport convertedReport = Assert.IsType<PdfRtfConversionReport>(result.Report);
        Assert.False(result.HasLoss);
        Assert.Null(result.Report.OutputPath);
        Assert.Contains(converted.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");

        using var sync = new MemoryStream();
        OfficeConversionReport syncReport = logicalSource
            ? logical.SaveAsRtf(sync, CreateImportOptions())
            : opened.SaveAsRtf(sync, CreateImportOptions());
        Assert.False(syncReport.HasLoss);
        Assert.Null(syncReport.OutputPath);
        Assert.IsType<PdfRtfConversionReport>(syncReport);
        using RtfDocument syncRtf = RtfDocument.ParseResult(Encoding.UTF8.GetString(sync.ToArray())).Document;
        Assert.Contains(syncRtf.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        Assert.True(sync.CanWrite);

        using var asyncOutput = new MemoryStream();
        OfficeConversionReport asyncReport = logicalSource
            ? await logical.SaveAsRtfAsync(asyncOutput, CreateImportOptions())
            : await opened.SaveAsRtfAsync(asyncOutput, CreateImportOptions());
        Assert.False(asyncReport.HasLoss);
        Assert.Null(asyncReport.OutputPath);
        Assert.IsType<PdfRtfConversionReport>(asyncReport);
        using RtfDocument asyncRtf = RtfDocument.ParseResult(Encoding.UTF8.GetString(asyncOutput.ToArray())).Document;
        Assert.Contains(asyncRtf.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");
        Assert.True(asyncOutput.CanWrite);

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-rtf-report-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string syncPath = Path.Combine(directory, "sync.rtf");
            OfficeConversionReport syncFile = logicalSource
                ? logical.SaveAsRtf(Path.Combine(directory, ".", "sync.rtf"), CreateImportOptions())
                : opened.SaveAsRtf(Path.Combine(directory, ".", "sync.rtf"), CreateImportOptions());
            Assert.Equal(Path.GetFullPath(syncPath), syncFile.OutputPath);
            Assert.Equal(convertedReport.Warnings.Select(w => (w.Code, w.LossKind)),
                Assert.IsType<PdfRtfConversionReport>(syncFile).Warnings.Select(w => (w.Code, w.LossKind)));
            using RtfDocument savedSync = RtfDocument.LoadResult(syncPath).Document;
            Assert.Contains(savedSync.Paragraphs, paragraph => paragraph.ToPlainText() == "Clinical Summary");

            string asyncPath = Path.Combine(directory, "async.rtf");
            OfficeConversionReport asyncFile = logicalSource
                ? await logical.SaveAsRtfAsync(asyncPath, CreateImportOptions())
                : await opened.SaveAsRtfAsync(asyncPath, CreateImportOptions());
            Assert.Equal(Path.GetFullPath(asyncPath), asyncFile.OutputPath);
            Assert.Equal(convertedReport.Diagnostics.Select(d => (d.Code, d.LossKind)),
                Assert.IsType<PdfRtfConversionReport>(asyncFile).Diagnostics.Select(d => (d.Code, d.LossKind)));
            using RtfDocument savedAsync = RtfDocument.LoadResult(asyncPath).Document;
            Assert.Contains(savedAsync.Paragraphs, paragraph => paragraph.ToPlainText() == "Second page body.");
        } finally {
            Directory.Delete(directory, recursive: true);
        }
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
        PdfCore.PdfDocumentReadResult logical = LoadSemanticPdf(pdf);
        Assert.NotEmpty(logical.Tables);

        logical.ToRtf(null, out OfficeConversionReport report);
        PdfRtfConversionReport result = (PdfRtfConversionReport)report;

        Assert.True(result.HasLoss);
        PdfCore.PdfConversionWarning warning = Assert.Single(
            result.Warnings,
            item => item.Code == "TABLES_NOT_IMPORTED");
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Equal(logical.Tables.Count.ToString(), warning.Details["count"]);

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-rtf-loss-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "table.rtf");
            PdfRtfConversionReport saved = Assert.IsType<PdfRtfConversionReport>(logical.SaveAsRtf(path));
            Assert.Equal(Path.GetFullPath(path), saved.OutputPath);
            Assert.True(saved.HasLoss);
            PdfCore.PdfConversionWarning savedWarning = Assert.Single(saved.Warnings, item => item.Code == "TABLES_NOT_IMPORTED");
            Assert.Equal(warning.LossKind, savedWarning.LossKind);
            Assert.Equal(warning.Details["count"], savedWarning.Details["count"]);
            Assert.Contains(saved.Diagnostics, diagnostic => diagnostic.Code == "TABLES_NOT_IMPORTED"
                && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => saved.ThrowIfLoss());
            Assert.Same(saved, error.Report);
            Assert.StartsWith(@"{\rtf1", File.ReadAllText(path));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task PdfToRtf_StrictLossRejectionLeavesCallerStreamUntouched() {
        using PdfCore.PdfDocument source = PdfCore.PdfDocument.Create().KeyValueTable(new[] {
            PdfCore.PdfKeyValueRow.Text("Invoice", "INV-001"),
            PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
            PdfCore.PdfKeyValueRow.Text("Due", "2026-06-30")
        });
        using var output = new MemoryStream(new byte[] { 1, 2, 3 }, writable: true);
        OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() =>
            source.SaveAsRtfAsync(output, new PdfRtfExportOptions { FailOnLoss = true }));
        Assert.IsType<PdfRtfConversionReport>(error.Report);
        Assert.Contains(error.Report.Diagnostics, diagnostic => diagnostic.Code == "TABLES_NOT_IMPORTED");
        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
        Assert.Equal(0, output.Position);
        Assert.True(output.CanWrite);
    }

    private static PdfRtfExportOptions CreateImportOptions() => new PdfRtfExportOptions();

    private static PdfCore.PdfDocumentReadResult LoadSemanticPdf(byte[] pdf) =>
        PdfCore.PdfDocumentReadResult.Load(pdf, CreateLayoutOptions());

    private static PdfCore.PdfTextLayoutOptions CreateLayoutOptions() => new PdfCore.PdfTextLayoutOptions {
        ForceSingleColumn = true
    };

    private static byte[] CreateSemanticPdf() =>
        PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageSetup = {
                    Size = new global::OfficeIMO.Pdf.PdfPageSize(global::OfficeIMO.OfficeLength.Points(360), global::OfficeIMO.OfficeLength.Points(360)),
                    Margins = { Left = global::OfficeIMO.OfficeLength.Points(36), Right = global::OfficeIMO.OfficeLength.Points(36), Top = global::OfficeIMO.OfficeLength.Points(36), Bottom = global::OfficeIMO.OfficeLength.Points(36) }
                },
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Meta(title: "PDF Import Title", author: "OfficeIMO", subject: "PDF to RTF", keywords: "pdf,rtf")
            .H1("Clinical Summary")
            .Paragraph(p => p.AddText("This semantic paragraph should become one imported RTF paragraph."))
            .Bullets(new[] { "First bullet", "Second bullet" })
            .Numbered(new[] { "First numbered", "Second numbered" }, startNumber: 3)
            .PageBreak()
            .Paragraph(p => p.AddText("Second page body."))
            .ToBytes();
}
