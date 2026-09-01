using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class PowerPointPdfEditableDefaultContracts {
    [Fact]
    public void DefaultEditableImportOmitsInvisibleTextAndReportsLoss() {
        byte[] pdf = BuildSingleStreamPdf(
            "BT /F1 12 Tf 72 720 Td (Visible text) Tj 3 Tr 0 -24 Td (Hidden OCR text) Tj 0 Tr ET");

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult();

        Assert.Equal(PdfPowerPointImportMode.EditableContent, result.Report.Mode);
        PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
        Assert.True(page.OmittedTextCount >= 1);
        Assert.True(result.HasLoss);
        Assert.Contains(result.Warnings, warning =>
            warning.Code == "PdfTextNotReconstructed" &&
            warning.Details.TryGetValue("Disposition", out string? disposition) &&
            disposition == "Omitted");
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        string[] text = package.PresentationPart!.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Text>())
            .Select(value => value.Text ?? string.Empty)
            .ToArray();
        Assert.Contains(text, value => value.Contains("Visible text", StringComparison.Ordinal));
        Assert.DoesNotContain(text, value => value.Contains("Hidden OCR text", StringComparison.Ordinal));
    }

    [Fact]
    public void DefaultEditableImportDoesNotExposeInvisibleTableCellText() {
        byte[] pdf = BuildSingleStreamPdf(string.Join("\n", new[] {
            "BT /F1 10 Tf",
            "50 700 Td (Name) Tj 100 0 Td (Value) Tj 100 0 Td (Status) Tj",
            "-200 -20 Td (Alpha) Tj 100 0 Td (42) Tj 100 0 Td (Ready) Tj",
            "-200 -20 Td (Beta) Tj 100 0 Td 3 Tr (Hidden table OCR) Tj 0 Tr 100 0 Td (Pending) Tj",
            "-200 -20 Td (Gamma) Tj 100 0 Td (84) Tj 100 0 Td (Done) Tj",
            "ET"
        }));

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult();

        Assert.Equal(PdfPowerPointImportMode.EditableContent, result.Report.Mode);
        Assert.True(result.HasLoss);
        Assert.Contains(result.Warnings, warning => warning.Code == "PdfTextNotReconstructed");
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        string[] text = package.PresentationPart!.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Text>())
            .Select(value => value.Text ?? string.Empty)
            .ToArray();
        Assert.Contains(text, value => value.Contains("Alpha", StringComparison.Ordinal));
        Assert.DoesNotContain(text, value => value.Contains("Hidden table OCR", StringComparison.Ordinal));
    }

    [Fact]
    public void DefaultEditableImportPreservesTableTextInsideTightProducerClip() {
        byte[] pdf = BuildSingleStreamPdf(string.Join("\n", new[] {
            "BT /F1 10 Tf",
            "50 700 Td (Name) Tj 100 0 Td (Value) Tj 100 0 Td (Status) Tj",
            "-200 -20 Td (Alpha) Tj 100 0 Td (42) Tj 100 0 Td (Ready) Tj",
            "ET",
            "q 45 655 310 14 re W n",
            "BT /F1 10 Tf 50 660 Td (Beta) Tj 100 0 Td (64) Tj 100 0 Td (Pending) Tj ET",
            "Q",
            "BT /F1 10 Tf 50 640 Td (Gamma) Tj 100 0 Td (84) Tj 100 0 Td (Done) Tj ET"
        }));

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult();

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.NotEmpty(package.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
        string[] text = package.PresentationPart.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Text>())
            .Select(value => value.Text ?? string.Empty)
            .ToArray();
        Assert.Contains(text, value => value.Contains("Beta", StringComparison.Ordinal));
        Assert.Contains(text, value => value.Contains("64", StringComparison.Ordinal));
    }

    [Fact]
    public void DefaultEditableImportDoesNotExposePartiallyClippedTableCellText() {
        byte[] pdf = BuildSingleStreamPdf(string.Join("\n", new[] {
            "BT /F1 10 Tf",
            "50 700 Td (Name) Tj 100 0 Td (Value) Tj 100 0 Td (Status) Tj",
            "-200 -20 Td (Alpha) Tj 100 0 Td (42) Tj 100 0 Td (Ready) Tj",
            "ET",
            "q 0 0 175 792 re W n",
            "BT /F1 10 Tf 50 660 Td (Beta) Tj 100 0 Td (Clipped table text) Tj 100 0 Td (Pending) Tj ET",
            "Q",
            "BT /F1 10 Tf 50 640 Td (Gamma) Tj 100 0 Td (84) Tj 100 0 Td (Done) Tj ET"
        }));

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult();

        Assert.True(result.HasLoss);
        Assert.Contains(result.Warnings, warning => warning.Code == "PdfTextNotReconstructed");
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        string[] text = package.PresentationPart!.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Text>())
            .Select(value => value.Text ?? string.Empty)
            .ToArray();
        Assert.Contains(text, value => value.Contains("Alpha", StringComparison.Ordinal));
        Assert.DoesNotContain(text, value => value.Contains("Clipped table text", StringComparison.Ordinal));
    }

    [Fact]
    public void DefaultEditableImportTreatsOmittedInteractiveContentAsLoss() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .H1("Approval")
            .TextField("Decision", width: 140, value: "Ready")
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult();

        Assert.True(result.Report.HasOmittedPageContent);
        Assert.True(result.HasLoss);
        Assert.Contains(result.Warnings, warning =>
            warning.Code == "PdfFormsNotReconstructed" &&
            warning.Details.TryGetValue("Disposition", out string? disposition) &&
            disposition == "Omitted");
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        result.Value.Dispose();
    }

    [Fact]
    public void LogicalPdfAutoProfileResolvesToEditableTablesForNullAndDefaultOptions() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Metric", "Value", "Status" },
                new[] { "Ready", "Yes", "Current" },
                new[] { "Loss", "Reported", "Current" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 100, 100, 120 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(
            pdf,
            new PdfCore.PdfTextLayoutOptions { ForceSingleColumn = true });

        PdfPowerPointConversionResult implicitResult = logical.ToPowerPointPresentationResult();
        PdfPowerPointConversionResult explicitDefaultResult = logical.ToPowerPointPresentationResult(
            new PdfPowerPointImportOptions());

        Assert.Equal(PdfPowerPointImportMode.EditableTables, implicitResult.Report.Mode);
        Assert.Equal(PdfPowerPointImportMode.EditableTables, explicitDefaultResult.Report.Mode);
        Assert.Single(implicitResult.Report.TableEntries);
        Assert.Single(explicitDefaultResult.Report.TableEntries);
        implicitResult.Value.Dispose();
        explicitDefaultResult.Value.Dispose();
    }

    private static byte[] BuildSingleStreamPdf(string streamContent) {
        streamContent = streamContent.TrimEnd('\n');
        int streamLength = Encoding.ASCII.GetByteCount(streamContent);
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {streamLength} >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }
}
