using System.Text.Json;
using System.Text.RegularExpressions;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReverseConversionScorecardTests {
    [Fact]
    public void Scorecard_ExecutesFiftySixCrossProducerReverseConversions() {
        string repositoryRoot = FindRepositoryRoot();
        using JsonDocument scorecard = JsonDocument.Parse(File.ReadAllBytes(
            Path.Combine(repositoryRoot, "Docs", "pdf-reverse-conversion-scorecard.json")));
        JsonElement root = scorecard.RootElement;
        JsonElement[] producers = root.GetProperty("producers").EnumerateArray().ToArray();
        JsonElement[] routes = root.GetProperty("routes").EnumerateArray().ToArray();
        string[] routeIds = routes
            .Select(static route => route.GetProperty("id").GetString()!)
            .ToArray();

        Assert.Equal(7, producers.Length);
        Assert.Equal(new[] {
            "pdf-to-docx", "pdf-to-html", "pdf-to-xlsx", "pdf-to-pptx",
            "pdf-to-odt", "pdf-to-ods", "pdf-to-odp", "pdf-to-png"
        }, routeIds);
        Assert.Equal(producers.Length * routeIds.Length, root.GetProperty("matrixCases").GetInt32());

        var executed = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement producer in producers) {
            string producerId = producer.GetProperty("id").GetString()!;
            string sourcePath = Path.Combine(repositoryRoot, producer.GetProperty("path").GetString()!.Replace('/', Path.DirectorySeparatorChar));
            Assert.True(File.Exists(sourcePath), "Missing scorecard fixture: " + sourcePath);
            byte[] source = File.ReadAllBytes(sourcePath);
            PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(source);
            Assert.NotEmpty(logical.Pages);
            HashSet<string> sourceTokens = GetTokens(string.Join(" ", logical.Pages.SelectMany(static page => page.TextBlocks).Select(static block => block.Text)));
            Assert.True(sourceTokens.Count > 0,
                "No source tokens were extracted for scorecard producer '" + producerId + "'. Page artifacts: " +
                string.Join(", ", logical.Pages.Select(static page =>
                    page.PageNumber + ": runs=" + page.Analysis.DecodedRuns.Count +
                    ", words=" + page.Analysis.Words.Count +
                    ", lines=" + page.Analysis.Lines.Count +
                    ", blocks=" + page.TextBlocks.Count)));
            bool expectedTables = producer.GetProperty("expectedTables").GetBoolean();
            if (expectedTables) Assert.NotEmpty(logical.Tables);

            foreach (JsonElement routeConfiguration in routes) {
                string route = routeConfiguration.GetProperty("id").GetString()!;
                Assert.True(executed.Add(producerId + ":" + route));
                ExecuteAndReopen(routeConfiguration, source, logical, sourceTokens, expectedTables);
            }
        }

        Assert.Equal(56, executed.Count);
    }

    [Fact]
    public async System.Threading.Tasks.Task Scorecard_ExecutesScannedMixedEncryptedAndMalformedEvidence() {
        string repositoryRoot = FindRepositoryRoot();
        using JsonDocument scorecard = JsonDocument.Parse(File.ReadAllBytes(
            Path.Combine(repositoryRoot, "Docs", "pdf-reverse-conversion-scorecard.json")));
        JsonElement[] stressCases = scorecard.RootElement.GetProperty("stressCases").EnumerateArray().ToArray();
        Assert.Equal(new[] { "scanned", "mixed-content", "encrypted", "malformed" },
            stressCases.Select(static item => item.GetProperty("kind").GetString()).ToArray());
        Assert.All(stressCases, item => Assert.NotEmpty(item.GetProperty("requiredEvidence").EnumerateArray()));

        byte[] scanned = PdfCore.PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 220, 90, alternativeText: "Scanned source")
            .ToBytes();
        var scannedProvider = new ScorecardOcrProvider(request => Result(new[] {
            OcrAt(request, "Scanned", 36, 120, 52, 14),
            OcrAt(request, "invoice", 96, 120, 44, 14)
        }));
        PdfOcrMergeResult scannedResult = await PdfCore.PdfDocument.Load(scanned).ReadWithOcrAsync(scannedProvider);
        Assert.Empty(scannedResult.NativeDocument.TextBlocks);
        Assert.Equal(2, scannedResult.AcceptedWordCount);
        Assert.All(scannedResult.Document.TextBlocks, block => Assert.Equal(PdfCore.PdfLogicalContentSourceKind.Ocr, block.SourceKind));
        using (OfficeIMO.Word.WordDocument word = scannedResult.Document.ToWordDocument()) {
            using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(word.ToBytes()), false);
            Assert.Contains("Scanned invoice", package.MainDocumentPart!.Document.InnerText, StringComparison.Ordinal);
        }

        byte[] mixed = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Native retained"))
            .Image(PdfPngTestImages.CreateRgbPng(120, 140, 160), 100, 40, alternativeText: "Mixed source")
            .ToBytes();
        PdfCore.PdfSelectionQuad nativeQuad = PdfCore.PdfPageInteractionMap.Create(mixed, 1).TextRegions[0].Quad;
        var mixedProvider = new ScorecardOcrProvider(request => Result(new[] {
            OcrAt(request, "duplicate", nativeQuad.Left, nativeQuad.Top, nativeQuad.Width, nativeQuad.Height),
            OcrAt(request, "OCR-retained", 36, 260, 74, 14)
        }));
        PdfOcrMergeResult mixedResult = await PdfCore.PdfDocument.Load(mixed).ReadWithOcrAsync(mixedProvider);
        Assert.Equal(1, mixedResult.Pages[0].RejectedNativeOverlapCount);
        Assert.Contains(mixedResult.Document.TextBlocks, block => block.SourceKind == PdfCore.PdfLogicalContentSourceKind.Native && block.Text.Contains("Native retained", StringComparison.Ordinal));
        Assert.Contains(mixedResult.Document.TextBlocks, block => block.SourceKind == PdfCore.PdfLogicalContentSourceKind.Ocr && block.Text.Contains("OCR-retained", StringComparison.Ordinal));

        byte[] encrypted = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Credentialed conversion"))
            .ToBytes();
        Assert.Throws<PdfCore.PdfPasswordRequiredException>(() => PdfCore.PdfDocumentReadResult.Load(encrypted));
        Assert.Throws<PdfCore.PdfInvalidPasswordException>(() => PdfCore.PdfDocumentReadResult.Load(encrypted, null, new PdfCore.PdfLoadOptions { Password = "wrong" }));
        PdfCore.PdfDocumentReadResult decrypted = PdfCore.PdfDocumentReadResult.Load(encrypted, null, new PdfCore.PdfLoadOptions { Password = "open" });
        Assert.Contains(decrypted.TextBlocks, block => block.Text.Contains("Credentialed conversion", StringComparison.Ordinal));
        Assert.Contains("Credentialed", decrypted.ToHtml(), StringComparison.Ordinal);

        JsonElement malformedCase = stressCases.Single(static item => item.GetProperty("kind").GetString() == "malformed");
        string malformedPath = Path.Combine(repositoryRoot, malformedCase.GetProperty("path").GetString()!.Replace('/', Path.DirectorySeparatorChar));
        byte[] malformed = File.ReadAllBytes(malformedPath);
        PdfCore.PdfDocumentReadResult recovered = PdfCore.PdfDocumentReadResult.Load(malformed);
        Assert.NotEmpty(recovered.Pages);
        Assert.Contains("<!doctype html", recovered.ToHtml(), StringComparison.OrdinalIgnoreCase);
        byte[] recoveredPng = PdfCore.PdfPageImageRenderer.RenderPageAsPng(malformed);
        AssertPng(recoveredPng);
    }

    private static void ExecuteAndReopen(
        JsonElement routeConfiguration,
        byte[] source,
        PdfCore.PdfDocumentReadResult logical,
        HashSet<string> sourceTokens,
        bool expectedTables) {
        string route = routeConfiguration.GetProperty("id").GetString()!;
        switch (route) {
            case "pdf-to-docx":
                using (OfficeIMO.Word.WordDocument document = logical.ToWordDocument()) {
                    using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(document.ToBytes()), false);
                    DocumentFormat.OpenXml.Wordprocessing.Body body = Assert.IsType<DocumentFormat.OpenXml.Wordprocessing.Body>(package.MainDocumentPart?.Document?.Body);
                    AssertTokenRecall(sourceTokens, string.Join(" ", body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(static text => text.Text)), routeConfiguration.GetProperty("minimumTokenRecall").GetDouble(), route);
                    if (expectedTables) Assert.NotEmpty(body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>());
                }
                return;
            case "pdf-to-html":
                string html = logical.ToHtml(new PdfToHtmlOptions { Profile = PdfHtmlProfile.Semantic });
                Assert.Contains("<!doctype html", html, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(logical.Pages.Count, Regex.Matches(html, "class=\"pdf-page\"[^>]*data-page-number=", RegexOptions.CultureInvariant).Count);
                AssertTokenRecall(sourceTokens, Regex.Replace(html, "<[^>]+>", " "), routeConfiguration.GetProperty("minimumTokenRecall").GetDouble(), route);
                return;
            case "pdf-to-xlsx":
                using (var stream = new MemoryStream()) {
                    PdfExcelTableImportReport report = logical.SaveTablesAsExcel(stream).RequireSuccess().Report!;
                    byte[] workbook = stream.ToArray();
                    if (expectedTables) Assert.NotEmpty(report.Entries);
                    HashSet<string>? tableTokens = null;
                    if (report.Entries.Count > 0) {
                        tableTokens = GetTokens(string.Join(" ", logical.Tables.SelectMany(static table => {
                            PdfCore.PdfLogicalTableData data = PdfCore.PdfLogicalTableAnalysis.Extract(table);
                            return data.Columns.Concat(data.Rows.SelectMany(static row => row));
                        })));
                    } else {
                        Assert.True(report.HasOmittedPageContent, "A table-only XLSX projection must report omitted non-table page content.");
                    }
                    using SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(workbook), false);
                    Assert.NotNull(package.WorkbookPart?.Workbook);
                    if (report.Entries.Count > 0) {
                        Assert.NotEmpty(package.WorkbookPart!.WorksheetParts.SelectMany(static worksheet => worksheet.TableDefinitionParts));
                        AssertTokenRecall(tableTokens!, GetSpreadsheetText(package), routeConfiguration.GetProperty("minimumTableTokenRecall").GetDouble(), route);
                    }
                }
                return;
            case "pdf-to-pptx":
                PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(source)
                    .ToPowerPointPresentationResult(PdfToPowerPointOptions.CreateEditableContent());
                using (result.Value) {
                    using var stream = new MemoryStream();
                    result.Value.Save(stream);
                    using PresentationDocument package = PresentationDocument.Open(new MemoryStream(stream.ToArray()), false);
                    Assert.Equal(logical.Pages.Count, package.PresentationPart!.SlideParts.Count());
                    string presentationText = string.Join(" ", package.PresentationPart.SlideParts.SelectMany(static slide => slide.Slide.Descendants<A.Text>()).Select(static text => text.Text));
                    AssertTokenRecall(sourceTokens, presentationText, routeConfiguration.GetProperty("minimumTokenRecall").GetDouble(), route);
                    if (expectedTables) Assert.NotEmpty(package.PresentationPart.SlideParts.SelectMany(static slide => slide.Slide.Descendants<A.Table>()));
                }
                return;
            case "pdf-to-odt": {
                PdfOdtConversionResult odtResult = logical.ToOdtDocumentResult();
                byte[] artifact = odtResult.Value.ToBytes();
                OdtDocument reopened = OdtDocument.Load(new MemoryStream(artifact));
                Assert.NotEmpty(reopened.Paragraphs);
                AssertTokenRecall(sourceTokens, ReadOpenDocumentText(artifact), routeConfiguration.GetProperty("minimumTokenRecall").GetDouble(), route);
                return;
            }
            case "pdf-to-ods": {
                PdfOdsConversionResult odsResult = logical.ToOdsDocumentResult();
                byte[] artifact = odsResult.Value.ToBytes();
                OdsDocument reopened = OdsDocument.Load(new MemoryStream(artifact));
                Assert.NotEmpty(reopened.Sheets);
                if (expectedTables) {
                    HashSet<string> tableTokens = GetTokens(string.Join(" ", logical.Tables.SelectMany(static table => {
                        PdfCore.PdfLogicalTableData data = PdfCore.PdfLogicalTableAnalysis.Extract(table);
                        return data.Columns.Concat(data.Rows.SelectMany(static row => row));
                    })));
                    AssertTokenRecall(tableTokens, ReadOpenDocumentText(artifact), routeConfiguration.GetProperty("minimumTableTokenRecall").GetDouble(), route);
                } else {
                    Assert.True(odsResult.Report.PdfReport.HasOmittedPageContent, "A table-only ODS projection must report omitted non-table page content.");
                }
                return;
            }
            case "pdf-to-odp": {
                PdfOdpConversionResult odpResult = logical.ToOdpPresentationResult(PdfToPowerPointOptions.CreateEditableContent());
                byte[] artifact = odpResult.Value.ToBytes();
                OdpPresentation reopened = OdpPresentation.Load(new MemoryStream(artifact));
                Assert.Equal(logical.Pages.Count, reopened.Slides.Count);
                AssertTokenRecall(sourceTokens, ReadOpenDocumentText(artifact), routeConfiguration.GetProperty("minimumTokenRecall").GetDouble(), route);
                return;
            }
            case "pdf-to-png": {
                var options = new PdfCore.PdfPageRenderOptions {
                    Format = PdfCore.PdfPageRenderFormat.Png,
                    Dpi = routeConfiguration.GetProperty("dpi").GetDouble(),
                    ContinueOnError = false
                };
                IReadOnlyList<PdfCore.PdfPageRenderResult> first = PdfCore.PdfPageImageRenderer.RenderPages(source, options: options);
                IReadOnlyList<PdfCore.PdfPageRenderResult> second = PdfCore.PdfPageImageRenderer.RenderPages(source, options: options);
                Assert.Equal(logical.Pages.Count, first.Count);
                Assert.Equal(first.Count, second.Count);
                for (int index = 0; index < first.Count; index++) {
                    Assert.True(first[index].Succeeded);
                    Assert.True(first[index].Width > 0 && first[index].Height > 0);
                    AssertPng(first[index].Bytes!);
                    Assert.Equal(first[index].Bytes, second[index].Bytes);
                }
                return;
            }
            default:
                throw new InvalidOperationException("Unknown reverse-conversion route: " + route);
        }
    }

    private static HashSet<string> GetTokens(string value) => Regex.Matches(value.ToLowerInvariant(), @"[\p{L}\p{N}]{4,}", RegexOptions.CultureInvariant)
        .Cast<Match>()
        .Select(static match => match.Value)
        .ToHashSet(StringComparer.Ordinal);

    private static void AssertTokenRecall(HashSet<string> expectedTokens, string actualText, double minimumRecall, string route) {
        Assert.NotEmpty(expectedTokens);
        HashSet<string> actualTokens = GetTokens(actualText);
        int retained = expectedTokens.Count(actualTokens.Contains);
        double recall = (double)retained / expectedTokens.Count;
        Assert.True(recall >= minimumRecall, route + " retained " + retained + "/" + expectedTokens.Count + " source tokens (" + recall.ToString("P1", System.Globalization.CultureInfo.InvariantCulture) + "); expected at least " + minimumRecall.ToString("P1", System.Globalization.CultureInfo.InvariantCulture) + ".");
    }

    private static string GetSpreadsheetText(SpreadsheetDocument package) {
        string[] sharedStrings = package.WorkbookPart?.SharedStringTablePart?.SharedStringTable?
            .Elements<S.SharedStringItem>()
            .Select(static item => item.InnerText)
            .ToArray() ?? Array.Empty<string>();
        var values = new List<string>();
        foreach (S.Cell cell in package.WorkbookPart!.WorksheetParts.SelectMany(static worksheet => worksheet.Worksheet.Descendants<S.Cell>())) {
            if (cell.DataType?.Value == S.CellValues.SharedString &&
                int.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int sharedIndex) &&
                sharedIndex >= 0 && sharedIndex < sharedStrings.Length) {
                values.Add(sharedStrings[sharedIndex]);
            } else if (cell.InlineString is not null) {
                values.Add(cell.InlineString.InnerText);
            } else if (cell.CellValue is not null) {
                values.Add(cell.CellValue.Text);
            }
        }
        return string.Join(" ", values);
    }

    private static string ReadOpenDocumentText(byte[] artifact) {
        using var archive = new ZipArchive(new MemoryStream(artifact), ZipArchiveMode.Read, leaveOpen: false);
        ZipArchiveEntry content = archive.GetEntry("content.xml") ?? throw new InvalidDataException("OpenDocument package did not contain content.xml.");
        using StreamReader reader = new StreamReader(content.Open());
        return Regex.Replace(reader.ReadToEnd(), "<[^>]+>", " ");
    }

    private static void AssertPng(byte[] artifact) {
        Assert.True(artifact.Length > 24);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, artifact.Take(8).ToArray());
        Assert.Equal("IHDR", System.Text.Encoding.ASCII.GetString(artifact, 12, 4));
    }

    private static OcrTextSpan OcrAt(
        OcrRequest request,
        string text,
        double x,
        double y,
        double width,
        double height) =>
        new OcrTextSpan {
            Level = OcrTextSpanLevel.Word,
            Text = text,
            Confidence = 0.98D,
            CoordinateUnit = OcrCoordinateUnit.Points,
            Region = new OcrRegion { X = x, Y = y, Width = width, Height = height }
        };

    private static OcrResult Result(IEnumerable<OcrTextSpan> spans) => new OcrResult { Spans = spans.ToArray() };

    private sealed class ScorecardOcrProvider : IOcrEngine {
        private readonly Func<OcrRequest, OcrResult> _response;
        internal ScorecardOcrProvider(Func<OcrRequest, OcrResult> response) { _response = response; }
        public string Id => "scorecard-fixture";
        public OcrEngineCapabilities Capabilities { get; } = new OcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/png" },
            SupportsWordSpans = true,
            SupportsConfidence = true
        };
        public System.Threading.Tasks.Task<OcrResult> RecognizeAsync(
            OcrRequest request,
            System.Threading.CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            return System.Threading.Tasks.Task.FromResult(_response(request));
        }
    }

    private static string FindRepositoryRoot() {
        string? current = AppContext.BaseDirectory;
        while (!string.IsNullOrWhiteSpace(current)) {
            if (File.Exists(Path.Combine(current, "OfficeIMO.sln"))) return current;
            current = Directory.GetParent(current)?.FullName;
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }
}
