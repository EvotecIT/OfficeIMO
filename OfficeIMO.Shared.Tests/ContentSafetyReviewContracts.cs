using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.ContentSafety;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.Rtf;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ContentSafetyReviewContracts {
    [Fact]
    public void ExcelSharedStringCleanupKeepsCellsDistinctAndParentRemovalDominatesExactEdits() {
        byte[] workbook = CreateSharedStringWorkbook("pay\u200Bload");
        OfficeContentSafetyReport report = ExcelDocument.InspectContentSafety(workbook, "shared.xlsx");
        OfficeContentSafetyFinding firstCell = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.HiddenContainer && item.Location.Contains("Cell(A1)", StringComparison.Ordinal));
        OfficeContentSafetyFinding firstUnicode = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.Location.Contains("Cell(A1)", StringComparison.Ordinal));
        OfficeContentSafetyFinding secondUnicode = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.Location.Contains("Cell(A2)", StringComparison.Ordinal));

        OfficeContentCleanupResult cleaned = ExcelDocument.RemoveSelectedContent(
            workbook,
            new OfficeContentCleanupSelection(new[] { firstCell.Id, firstUnicode.Id, secondUnicode.Id }),
            "shared.xlsx");

        Assert.DoesNotContain(cleaned.After.Findings, item => item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode);
        using var stream = new MemoryStream(cleaned.Output, writable: false);
        using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
        S.Cell[] cells = package.WorkbookPart!.WorksheetParts.Single(part =>
            part.Worksheet.Descendants<S.Cell>().Any()).Worksheet.Descendants<S.Cell>().ToArray();
        Assert.Equal(string.Empty, cells.Single(item => item.CellReference?.Value == "A1").InnerText);
        Assert.Equal("payload", cells.Single(item => item.CellReference?.Value == "A2").InnerText);
    }

    [Fact]
    public void ExcelDrawingInspectionIncludesFieldsHiddenGroupsAndExplicitContrast() {
        byte[] workbook = CreateDrawingWorkbook();

        OfficeContentSafetyReport report = ExcelDocument.InspectContentSafety(workbook, "drawing.xlsx");

        Assert.Contains(report.Findings, item => item.Location.Contains("DrawingField", StringComparison.Ordinal) &&
            item.Kind == OfficeContentConcealmentKind.HiddenContainer && item.TextPreview.Contains("field payload", StringComparison.Ordinal));
        Assert.Contains(report.Findings, item => item.Location.Contains("DrawingRun", StringComparison.Ordinal) &&
            item.Kind == OfficeContentConcealmentKind.LowContrastText && item.TextPreview.Contains("white drawing", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlFilterOpacityRequiresAValidTopLevelFunction() {
        const string html = "<html><body>" +
            "<p style='filter:url(&quot;#opacity(0)&quot;)'>url remains visible</p>" +
            "<p style='filter:opacity(-1)'>invalid remains visible</p>" +
            "<p style='filter:blur(1px) /* retained */ opacity(200%) opacity(0)'>valid hidden</p>" +
            "</body></html>";

        OfficeContentSafetyReport report = HtmlContentSafety.Inspect(html);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview.Contains("url remains", StringComparison.Ordinal));
        Assert.DoesNotContain(report.Findings, item => item.TextPreview.Contains("invalid remains", StringComparison.Ordinal));
        Assert.Contains(report.Findings, item => item.Kind == OfficeContentConcealmentKind.TransparentText &&
            item.TextPreview.Contains("valid hidden", StringComparison.Ordinal));
    }

    [Fact]
    public void EmptyHtmlFileCleanupPreservesOriginalEncodingAndBomBytes() {
        string input = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".html");
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".html");
        try {
            foreach (byte[] bytes in new[] {
                Join(new UTF8Encoding(true).GetPreamble(), new UTF8Encoding(false).GetBytes("<p>visible</p>")),
                Join(Encoding.Unicode.GetPreamble(), Encoding.Unicode.GetBytes("<p>visible</p>"))
            }) {
                File.WriteAllBytes(input, bytes);

                OfficeContentCleanupResult result = HtmlContentSafety.RemoveSelectedFile(
                    input,
                    output,
                    new OfficeContentCleanupSelection(Array.Empty<string>()));

                Assert.False(result.Changed);
                Assert.Equal(bytes, result.Output);
                Assert.Equal(bytes, File.ReadAllBytes(output));
            }
        } finally {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void HtmlMachineOnlyUnicodeCanBeCleanedWithoutDeletingTheCarrier() {
        const string html = "<html><body><script>const pay\u200Bload = 'keep';</script></body></html>";
        OfficeContentSafetyReport report = HtmlContentSafety.Inspect(html);
        OfficeContentSafetyFinding unicode = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.TextPreview == "\\u200B");

        OfficeContentCleanupResult cleaned = HtmlContentSafety.RemoveSelected(html, new OfficeContentCleanupSelection(new[] { unicode.Id }));
        string output = Encoding.UTF8.GetString(cleaned.Output);

        Assert.Contains("<script>", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("const payload", output, StringComparison.Ordinal);
        Assert.Contains(cleaned.After.Findings, item => item.Location.Contains("script", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void RtfHtmlUnicodeCanBeCleanedWithoutDeletingTheEncapsulation() {
        byte[] rtf = Encoding.ASCII.GetBytes(@"{\rtf1\ansi\fromhtml1{\*\htmltag <html><body>pay\u8203?load</body></html>}ordinary visible}");
        OfficeContentSafetyReport report = RtfDocument.InspectContentSafety(rtf);
        OfficeContentSafetyFinding unicode = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.Location.StartsWith("HtmlEncapsulation", StringComparison.Ordinal));

        OfficeContentCleanupResult cleaned = RtfDocument.RemoveSelectedContent(rtf, new OfficeContentCleanupSelection(new[] { unicode.Id }));

        RtfDocument reopened = RtfDocument.LoadResult(cleaned.Output).Document;
        Assert.NotNull(reopened.HtmlEncapsulation);
        Assert.Contains("payload", reopened.HtmlEncapsulation!.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void OpenDocumentAttributeUnicodeCanBeCleanedWithoutDeletingTheStoredValue() {
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("ordinary visible");
        byte[] bytes = InjectOdfHiddenText(document.ToBytes(), "pay\u200Bload");
        OfficeContentSafetyReport report = OdfDocument.InspectContentSafety(bytes);
        OfficeContentSafetyFinding unicode = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.Location.Contains("string-value", StringComparison.Ordinal));

        OfficeContentCleanupResult cleaned = OdfDocument.RemoveSelectedContent(bytes, new OfficeContentCleanupSelection(new[] { unicode.Id }));

        Assert.Contains(cleaned.After.Findings, item => item.Kind == OfficeContentConcealmentKind.HiddenByProperty && item.TextPreview.Contains("payload", StringComparison.Ordinal));
        Assert.DoesNotContain(cleaned.After.Findings, item => item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode);
    }

    [Fact]
    public void PdfPaintedLowContrastUnicodeCanBeCleanedWithoutDeletingTheSpan() {
        byte[] pdf = CreateLowContrastPdfWithSoftHyphen();
        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(pdf);
        OfficeContentSafetyFinding unicode = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.TextPreview == "\\u00AD");
        Assert.Equal(OfficeContentCleanupCapability.RemoveText, unicode.CleanupCapability);

        OfficeContentCleanupResult cleaned = PdfDocument.RemoveSelectedContent(pdf, new OfficeContentCleanupSelection(new[] { unicode.Id }));

        string text = string.Concat(PdfReadDocument.Open(cleaned.Output).Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text));
        Assert.Contains("payload", text, StringComparison.Ordinal);
        Assert.Contains(cleaned.After.Findings, item => item.Kind == OfficeContentConcealmentKind.LowContrastText);
    }

    private static byte[] CreateSharedStringWorkbook(string text) {
        using var stream = new MemoryStream();
        using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true)) {
            WorkbookPart workbookPart = package.AddWorkbookPart();
            workbookPart.Workbook = new S.Workbook(new S.Sheets());
            SharedStringTablePart stringsPart = workbookPart.AddNewPart<SharedStringTablePart>();
            stringsPart.SharedStringTable = new S.SharedStringTable(new S.SharedStringItem(new S.Text(text)));

            WorksheetPart visiblePart = workbookPart.AddNewPart<WorksheetPart>();
            visiblePart.Worksheet = new S.Worksheet(new S.SheetData());
            WorksheetPart hiddenPart = workbookPart.AddNewPart<WorksheetPart>();
            hiddenPart.Worksheet = new S.Worksheet(new S.SheetData(
                new S.Row(
                    new S.Cell { CellReference = "A1", DataType = S.CellValues.SharedString, CellValue = new S.CellValue("0") },
                    new S.Cell { CellReference = "A2", DataType = S.CellValues.SharedString, CellValue = new S.CellValue("0") }) { RowIndex = 1U }));
            workbookPart.Workbook.Sheets!.Append(
                new S.Sheet { Id = workbookPart.GetIdOfPart(visiblePart), SheetId = 1U, Name = "Visible" },
                new S.Sheet { Id = workbookPart.GetIdOfPart(hiddenPart), SheetId = 2U, Name = "Hidden", State = S.SheetStateValues.Hidden });
            visiblePart.Worksheet.Save();
            hiddenPart.Worksheet.Save();
            stringsPart.SharedStringTable.Save();
            workbookPart.Workbook.Save();
        }
        return stream.ToArray();
    }

    private static byte[] CreateDrawingWorkbook() {
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create()) {
            document.AddWorksheet("Drawing");
            bytes = document.ToBytes(ExcelFileFormat.Xlsx);
        }
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, true)) {
            WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
            DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            worksheetPart.Worksheet.Append(new S.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing(
                new Xdr.TwoCellAnchor(
                    MarkerFrom(0, 0),
                    MarkerTo(3, 3),
                    new Xdr.GroupShape(
                        new Xdr.NonVisualGroupShapeProperties(
                            new Xdr.NonVisualDrawingProperties { Id = 10U, Name = "Hidden group", Hidden = true },
                            new Xdr.NonVisualGroupShapeDrawingProperties()),
                        new Xdr.GroupShapeProperties(new A.TransformGroup(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 2_000_000L, Cy = 2_000_000L },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 2_000_000L, Cy = 2_000_000L })),
                        DrawingShape(11U, "Field shape", new A.Field(
                            new A.RunProperties(),
                            new A.Text("field payload")) { Id = "{00000000-0000-0000-0000-000000000001}", Type = "test" })),
                    new Xdr.ClientData()),
                new Xdr.TwoCellAnchor(
                    MarkerFrom(4, 0),
                    MarkerTo(7, 3),
                    DrawingShape(20U, "White shape", new A.Run(
                        new A.RunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" })),
                        new A.Text("white drawing")), "FFFFFF"),
                    new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }
        return stream.ToArray();
    }

    private static Xdr.Shape DrawingShape(uint id, string name, OpenXmlElement textElement, string? fill = null) => new Xdr.Shape(
        new Xdr.NonVisualShapeProperties(
            new Xdr.NonVisualDrawingProperties { Id = id, Name = name },
            new Xdr.NonVisualShapeDrawingProperties()),
        new Xdr.ShapeProperties(
            new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 1_000_000L, Cy = 1_000_000L }),
            new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle },
            fill == null ? null : new A.SolidFill(new A.RgbColorModelHex { Val = fill })),
        new Xdr.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(textElement)));

    private static Xdr.FromMarker MarkerFrom(int column, int row) => new Xdr.FromMarker(
        new Xdr.ColumnId(column.ToString()), new Xdr.ColumnOffset("0"),
        new Xdr.RowId(row.ToString()), new Xdr.RowOffset("0"));

    private static Xdr.ToMarker MarkerTo(int column, int row) => new Xdr.ToMarker(
        new Xdr.ColumnId(column.ToString()), new Xdr.ColumnOffset("0"),
        new Xdr.RowId(row.ToString()), new Xdr.RowOffset("0"));

    private static byte[] InjectOdfHiddenText(byte[] bytes, string value) {
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry("content.xml")!;
            XDocument content;
            using (Stream source = entry.Open()) content = XDocument.Load(source);
            XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            content.Descendants(office + "text").Single().Add(new XElement(text + "hidden-text",
                new XAttribute(text + "is-hidden", "true"),
                new XAttribute(text + "string-value", value)));
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("content.xml", CompressionLevel.Optimal);
            using Stream destination = replacement.Open();
            content.Save(destination, SaveOptions.DisableFormatting);
        }
        return stream.ToArray();
    }

    private static byte[] CreateLowContrastPdfWithSoftHyphen() {
        const string content = "BT /F1 12 Tf 0 Tr 1 1 1 rg 72 720 Td (pay\\255load) Tj ET\n";
        string[] objects = {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            "<< /Length " + Encoding.ASCII.GetByteCount(content) + " >>\nstream\n" + content + "endstream",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
        };
        var output = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int> { 0 };
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(Encoding.ASCII.GetByteCount(output.ToString()));
            output.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = Encoding.ASCII.GetByteCount(output.ToString());
        output.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        for (int index = 1; index < offsets.Count; index++) output.Append(offsets[index].ToString("D10")).Append(" 00000 n \n");
        output.Append("trailer\n<< /Size ").Append(objects.Length + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(output.ToString());
    }

    private static byte[] Join(byte[] left, byte[] right) {
        var output = new byte[left.Length + right.Length];
        Buffer.BlockCopy(left, 0, output, 0, left.Length);
        Buffer.BlockCopy(right, 0, output, left.Length, right.Length);
        return output;
    }
}
