using CodeGlyphX;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Drawing.CodeGlyphX.Tests;

// Each format's one-call placement is checked at the public boundary: the saved file reopens with the
// code at its documented size and alternative text, the stored symbol decodes, and the page OfficeIMO's
// PDF conversion produces draws it as vector content that still decodes.
public sealed class DocumentBarcodePlacementTests : IDisposable {
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "officeimo-codeglyphx-" + Guid.NewGuid().ToString("N"));

    public DocumentBarcodePlacementTests() => Directory.CreateDirectory(_directory);

    public void Dispose() {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, recursive: true);
    }

    [Fact]
    public void WordQrCodeReopensWithAltTextAndConvertsToScannablePdf() {
        const string payload = "https://evotec.xyz/officeimo";
        string path = Path.Combine(_directory, "qr.docx");
        using (WordDocument document = WordDocument.Create(path)) {
            document.AddParagraph("Scan to open the guide:");
            document.AddQrCode(payload);
            document.Save();
        }

        using WordDocument reopened = WordDocument.Load(path);
        WordImage image = Assert.Single(reopened.Images);
        Assert.Equal("QR Code: " + payload, image.AltText);
        AssertLength(OfficeLength.Inches(1.25), image.Width!.Value);
        AssertLength(image.Width!.Value, image.Height!.Value);
        Assert.Equal(payload, ScanSingle(RasterizeSvg(image.ToBytes()), SymbolFormat.QrCode));

        using var pdf = new MemoryStream();
        reopened.SaveAsPdf(pdf);
        AssertPdfPageDecodes(pdf.ToArray(), SymbolFormat.QrCode, payload);
    }

    [Fact]
    public void DenseWordQrCodeStaysVectorAndScannableInPdf() {
        // About 1,500 characters need tens of thousands of module-run commands.
        string payload = string.Concat(Enumerable.Range(0, 1500).Select(index => (char)('a' + index * 7 % 26)));
        using WordDocument document = WordDocument.Create();
        document.AddQrCode(payload, options => options.Width = OfficeLength.Inches(5));

        using var pdf = new MemoryStream();
        document.SaveAsPdf(pdf);

        AssertPdfPageDecodes(pdf.ToArray(), SymbolFormat.QrCode, payload);
    }

    [Fact]
    public void ExcelBarcodeReopensAtCellAndConvertsToScannablePdf() {
        const string payload = "ORDER-2026-0042";
        string path = Path.Combine(_directory, "code128.xlsx");
        using (ExcelDocument workbook = ExcelDocument.Create(path)) {
            ExcelSheet sheet = workbook.AddSheet("Labels");
            sheet.Cell(1, 1).SetValue("Order");
            sheet.AddBarcode("B2", SymbolFormat.Code128, payload, options => options.Width = OfficeLength.Millimeters(60));
            workbook.Save();
        }

        using ExcelDocument reopened = ExcelDocument.Load(path);
        ExcelImage image = Assert.Single(reopened.Sheets[0].Images);
        Assert.Equal("Code 128: " + payload, image.Description);
        Assert.Equal(2, image.RowIndex);
        Assert.Equal(2, image.ColumnIndex);
        AssertLength(OfficeLength.Millimeters(60), image.Width);
        Assert.Equal("image/svg+xml", image.ContentType);
        Assert.Equal(payload, ScanSingle(RasterizeSvg(image.ToBytes()), SymbolFormat.Code128));

        using var pdf = new MemoryStream();
        reopened.SaveAsPdf(pdf);
        AssertPdfPageDecodes(pdf.ToArray(), SymbolFormat.Code128, payload);
    }

    [Fact]
    public void PowerPointQrCodeFillsLayoutBoxCenteredAndConvertsToScannablePdf() {
        const string payload = "WIFI:T:WPA;S:Guest;P:welcome-2026;;";
        string path = Path.Combine(_directory, "qr.pptx");
        var box = new PowerPointLayoutBox(OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(4), OfficeLength.Inches(2));
        using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddQrCode(payload, box, options => options.AlternativeText = "Guest Wi-Fi");
            presentation.Save();
        }

        using PowerPointDocument reopened = PowerPointDocument.Load(path);
        PowerPointImage picture = Assert.Single(reopened.Slides[0].Images);
        Assert.Equal("Guest Wi-Fi", picture.AltText);
        AssertLength(OfficeLength.Inches(2), picture.Height);
        AssertLength(OfficeLength.Inches(2), picture.Width);
        AssertLength(OfficeLength.Inches(2), picture.Left);
        AssertLength(OfficeLength.Inches(1), picture.Top);
        Assert.Equal(payload, ScanSingle(RasterizeSvg(picture.GetImageBytes()), SymbolFormat.QrCode));

        using var pdf = new MemoryStream();
        reopened.SaveAsPdf(pdf);
        AssertPdfPageDecodes(pdf.ToArray(), SymbolFormat.QrCode, payload);
    }

    [Theory]
    [InlineData(SymbolFormat.QrCode, "https://evotec.xyz", "https://evotec.xyz")]
    [InlineData(SymbolFormat.DataMatrix, "SN-0042-PL", "SN-0042-PL")]
    [InlineData(SymbolFormat.Pdf417, "Boarding pass 0042 seat 12C", "Boarding pass 0042 seat 12C")]
    [InlineData(SymbolFormat.Code128, "INV-2026-11-14", "INV-2026-11-14")]
    [InlineData(SymbolFormat.Gs1Code128, "(01)09501101530003(17)261231", "010950110153000317261231")]
    [InlineData(SymbolFormat.Ean, "590123412345", "5901234123457")]
    [InlineData(SymbolFormat.UpcA, "03600029145", "036000291452")]
    public void PdfFlowDrawsEachSymbologyAsScannableVector(SymbolFormat format, string payload, string expected) {
        // A label-sized page keeps the scanned image small.
        byte[] bytes = SavePdf(document => document.Page(page => page
            .Size(OfficeLength.Inches(5), OfficeLength.Inches(4))
            .Margin(OfficeLength.Inches(0.25))
            .Content(content => content
                .AddParagraph("Label")
                .AddBarcode(format, payload))));

        using PdfDocument pdf = PdfDocument.Load(new MemoryStream(bytes));
        Assert.Empty(pdf.Images.Extract());
        string decoded = ScanSingle(pdf.Render.DisplayPage(1, new PdfPageDisplayOptions { Scale = 2D }).Bytes!, format);
        Assert.Equal(expected, NormalizeDecoded(decoded));
    }

    [Fact]
    public void PdfCanvasPlacesQrCodeAtRequestedPositionAndSize() {
        const string payload = "TICKET-7781";
        byte[] bytes = SavePdf(document => document.Page(page => page.Size(OfficeLength.Inches(3), OfficeLength.Inches(3)).Canvas(canvas =>
            canvas.QrCode(payload, OfficeLength.Millimeters(20), OfficeLength.Millimeters(30), options => {
                options.Width = OfficeLength.Millimeters(25);
                options.ErrorCorrection = QrErrorCorrectionLevel.H;
            }))));

        using PdfDocument pdf = PdfDocument.Load(new MemoryStream(bytes));
        Assert.Empty(pdf.Images.Extract());
        OfficeDrawing page = pdf.Render.Drawing(1);
        Assert.Equal(payload, ScanSingle(OfficeDrawingRasterRenderer.ToPng(page, 4D), SymbolFormat.QrCode));
    }

    [Fact]
    public void LinearBarcodeHonorsExplicitWidthAndHeightIncludingText() {
        using WordDocument document = WordDocument.Create();
        WordImage image = document.AddBarcode(SymbolFormat.Code128, "PALLET-0042", options => {
            options.Width = OfficeLength.Millimeters(70);
            options.Height = OfficeLength.Millimeters(20);
        });

        AssertLength(OfficeLength.Millimeters(70), image.Width!.Value);
        AssertLength(OfficeLength.Millimeters(20), image.Height!.Value, toleranceMillimeters: 0.4);
        Assert.Equal("PALLET-0042", ScanSingle(RasterizeSvg(image.ToBytes()), SymbolFormat.Code128));
    }

    [Fact]
    public void RetailBarcodePrintsTheCheckDigitUnderTheBars() {
        using WordDocument document = WordDocument.Create();
        WordImage image = document.AddBarcode(SymbolFormat.Ean, "590123412345");

        Assert.Contains(">5901234123457</text>", System.Text.Encoding.UTF8.GetString(image.ToBytes()));
    }

    [Theory]
    [InlineData("error-correction-on-linear")]
    [InlineData("text-on-matrix")]
    [InlineData("theme-color")]
    [InlineData("negative-quiet-zone")]
    public void InvalidOptionsAreRejectedBeforeTheDocumentChanges(string scenario) {
        using WordDocument document = WordDocument.Create();
        int paragraphs = document.Paragraphs.Count;

        Assert.ThrowsAny<ArgumentException>(() => {
            switch (scenario) {
                case "error-correction-on-linear":
                    document.AddBarcode(SymbolFormat.Code128, "A1", o => o.ErrorCorrection = QrErrorCorrectionLevel.H);
                    break;
                case "text-on-matrix":
                    document.AddBarcode(SymbolFormat.DataMatrix, "A1", o => o.HumanReadableText = true);
                    break;
                case "theme-color":
                    document.AddQrCode("A1", o => o.Color = OfficeColor.Accent1);
                    break;
                default:
                    document.AddQrCode("A1", o => o.QuietZone = -1);
                    break;
            }
        });
        Assert.Equal(paragraphs, document.Paragraphs.Count);
    }

    private byte[] SavePdf(Action<PdfDocumentBuilder> compose) {
        using var stream = new MemoryStream();
        using (PdfDocument document = PdfDocument.Create(compose)) document.Save(stream);
        return stream.ToArray();
    }

    private static void AssertPdfPageDecodes(byte[] bytes, SymbolFormat format, string payload) {
        using PdfDocument pdf = PdfDocument.Load(new MemoryStream(bytes));
        // The symbol stays vector: no image XObject stands in for it.
        Assert.Empty(pdf.Images.Extract());
        byte[] page = pdf.Render.DisplayPage(1, new PdfPageDisplayOptions { Scale = 2D }).Bytes!;
        Assert.Equal(payload, ScanSingle(page, format));
    }

    private static byte[] RasterizeSvg(byte[] svg) {
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        return OfficeDrawingRasterRenderer.ToPng(drawing!, 1D);
    }

    private static string ScanSingle(byte[] png, SymbolFormat format) {
        // A whole page is a large image; give the scanner room on slower runtimes such as .NET Framework.
        ScanResult result = SymbolScanner.Scan(png, new ScanOptions { Formats = new[] { format }, MaxSymbols = 1, TimeoutMilliseconds = 30000 });
        DetectedSymbol symbol = Assert.Single(result.Symbols);
        return symbol.Text;
    }

    // GS1 element strings decode with optional parentheses or a symbology identifier; compare the element data.
    private static string NormalizeDecoded(string value) =>
        value.StartsWith("]C1", StringComparison.Ordinal) ? value.Substring(3) : value.Replace("(", string.Empty).Replace(")", string.Empty).Replace("\u001d", string.Empty);

    private static void AssertLength(OfficeLength expected, OfficeLength actual, double toleranceMillimeters = 0.05) =>
        Assert.InRange(Math.Abs(expected.ToPoints() - actual.ToPoints()) * 25.4 / 72D, 0D, toleranceMillimeters);
}
