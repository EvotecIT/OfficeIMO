using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Tests.Pdf;
using System.IO.Compression;
using Xunit;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Tests;

public sealed partial class ReaderOcrCoreTests {
    [Fact]
    public async Task TreeOcrPreservesNestedRichResultsAndProjectsOnlyNewTextOnce() {
        var grandchild = CreateDocument(1);
        grandchild.Source.Path = "scan.png";
        grandchild.OcrCandidates[0].Location.Path = "scan.png";
        grandchild.Links = new[] { new OfficeDocumentLink { Id = "link", Uri = "https://example.test" } };
        grandchild.Forms = new[] { new OfficeDocumentFormField { Id = "form", Value = "native" } };
        var child = CreateDocument(0);
        child.Source.Path = "attachment.zip";
        child.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "scan.png", Document = grandchild } };
        var source = CreateDocument(0);
        source.Source.Path = "mail.eml";
        source.Blocks = new[] { new OfficeDocumentBlock { Id = "body", Text = "Native body", Location = new ReaderLocation { Path = "mail.eml" } } };
        source.Markdown = "Native body";
        source.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "attachment.zip", Document = child } };
        string before = source.ToJson();
        var engine = new DelegateOcrEngine("tree", (_, _) => Task.FromResult(new OcrResult { Text = "Scanned amount 42" }));

        var result = await source.ApplyOcrTreeAsync(engine);
        Assert.Equal(3, result.Report.DocumentCount);
        var recognition = Assert.Single(result.Recognitions);
        Assert.Equal("root/n0/n0", recognition.DocumentId);
        Assert.Equal("mail.eml!/attachment.zip!/scan.png", recognition.DocumentPath);
        var block = Assert.Single(result.Document.Blocks, block => block.Kind == "ocr-text");
        Assert.Equal("root/n0/n0/ocr-1-block", block.Id);
        Assert.Equal("mail.eml!/attachment.zip!/scan.png", block.Location.Path);
        Assert.Equal("Scanned amount 42", Assert.Single(result.Document.Chunks).Text);
        Assert.Single(result.Document.EnumerateContent(), item => item.Block?.Text == "Scanned amount 42" || item.Chunk?.Text == "Scanned amount 42");
        Assert.Equal("Native body\n\nScanned amount 42", result.Document.Markdown);
        var restored = OfficeDocumentReadResultJson.Deserialize(result.Document.ToJson());
        var scan = Assert.Single(Assert.Single(restored.NestedDocuments).Document.NestedDocuments).Document;
        Assert.Empty(scan.OcrCandidates);
        Assert.Single(scan.Links);
        Assert.Single(scan.Forms);
        Assert.Single(scan.Assets);
        Assert.Equal(before, source.ToJson());
    }

    [Theory]
    [InlineData("candidates")]
    [InlineData("bytes")]
    [InlineData("text")]
    public async Task TreeOcrLimitsAreSharedAcrossAttachments(string budget) {
        var source = CreateDocument(0);
        source.NestedDocuments = Enumerable.Range(0, 3).Select(index => new OfficeDocumentNestedResult {
            Path = "scan-" + index + ".png", Document = CreateDocument(1)
        }).ToArray();
        int calls = 0;
        var engine = new DelegateOcrEngine("tree", (_, _) => { calls++; return Task.FromResult(new OcrResult { Text = "ABCDEF" }); });
        var options = new OfficeDocumentOcrExecutionOptions();
        if (budget == "candidates") options.MaxCandidates = 1;
        if (budget == "bytes") options.MaxTotalInputBytes = 1;
        if (budget == "text") options.MaxTotalRecognizedCharacters = 5;

        var result = await source.ApplyOcrTreeAsync(engine, options);

        Assert.Equal(1, calls);
        Assert.Equal(3, result.Report.CandidateCount);
        Assert.Equal(1, result.Report.AttemptedCandidateCount);
        Assert.Equal(2, result.Report.SkippedCandidateCount);
        Assert.Equal(1, result.Report.InputBytes);
        Assert.Single(result.Document.Blocks);
        Assert.Equal(budget == "text" ? "ABCDE" : "ABCDEF", Assert.Single(result.Recognitions).Result.Text);
        Assert.All(result.Document.NestedDocuments.Skip(1), nested => Assert.Single(nested.Document.OcrCandidates));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Category == OfficeDocumentDiagnosticCategory.Limit);
        Assert.All(result.Diagnostics, diagnostic => Assert.StartsWith("scan.pdf!/scan-", diagnostic.Location!.Path));
    }

    [Fact]
    public async Task TotalSpanBudgetsBoundRetainedOutputAcrossCandidates() {
        var engine = new DelegateOcrEngine("spans", (_, _) => Task.FromResult(new OcrResult {
            Text = "text", Spans = Enumerable.Range(0, 3).Select(index => new OcrTextSpan {
                Text = "1234", Sequence = index, Level = OcrTextSpanLevel.Word
            }).ToArray()
        }));
        var result = await CreateDocument(2).ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxTotalSpans = 4, MaxTotalSpanCharacters = 5
        });
        Assert.Equal(2, result.Recognitions.Count);
        Assert.Equal(4, result.Report.WordSpanCount);
        Assert.Equal(5, result.Recognitions.Sum(item => item.Result.Spans.Sum(span => span.Text.Length)));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ocr-span-limit");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ocr-span-text-limit");
    }

    [Theory]
    [InlineData(1, 8)]
    [InlineData(100, 0)]
    public async Task TreeDepthAndDocumentLimitsRetainPendingContents(int documents, int depth) {
        var source = CreateDocument(0);
        source.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "scan", Document = CreateDocument(1) } };
        int calls = 0;
        var engine = new DelegateOcrEngine("tree", (_, _) => { calls++; return Task.FromResult(new OcrResult { Text = "text" }); });
        var result = await source.ApplyOcrTreeAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxDocuments = documents, MaxNestedDepth = depth
        });
        Assert.Equal(0, calls);
        Assert.Equal(1, result.Report.SkippedDocumentCount);
        Assert.Single(Assert.Single(result.Document.NestedDocuments).Document.OcrCandidates);
        Assert.Contains(result.Document.Diagnostics, diagnostic => diagnostic.Code == "ocr-nested-document-limit");
    }

    [Fact]
    public async Task PageOwnedCandidatesExecuteAfterTransportAndPayloadMaterialization() {
        var source = CreateDocument(1);
        source.Pages[0].Assets = source.Assets;
        source.Assets = Array.Empty<OfficeDocumentAsset>();
        source.OcrCandidates = Array.Empty<OfficeDocumentOcrCandidate>();
        byte[] payload = source.Pages[0].Assets[0].PayloadBytes!;
        source = OfficeDocumentReadResultJson.Deserialize(source.ToJson());
        source.Pages[0].Assets[0].PayloadBytes = payload; // Binary payloads are deliberately not part of Reader JSON.
        var engine = new DelegateOcrEngine("page", (_, _) => Task.FromResult(new OcrResult { Text = "page text" }));
        var result = await source.ApplyOcrAsync(engine);
        Assert.Equal(1, result.Report.CandidateCount);
        Assert.Single(result.Recognitions);
        Assert.Empty(result.Document.Pages[0].OcrCandidates);
        Assert.Equal("page text", Assert.Single(result.Document.Pages[0].Blocks).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RegisteredImagesInsideArchivesReachParentOcrEvidence(bool emailContainer) {
        byte[] image = PdfPngTestImages.CreateRgbPng(200, 200, 200);
        using var bytes = new MemoryStream();
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true)) {
            using Stream payload = zip.CreateEntry("scan.png").Open();
            payload.Write(image, 0, image.Length);
        }
        byte[] sourceBytes = bytes.ToArray();
        string name = "attachment.zip";
        if (emailContainer) {
            string mime = "MIME-Version: 1.0\r\nSubject: Scan\r\nContent-Type: multipart/mixed; boundary=probe\r\n\r\n"
                + "--probe\r\nContent-Type: text/plain\r\n\r\nNative body\r\n"
                + "--probe\r\nContent-Type: application/zip; name=attachment.zip\r\nContent-Disposition: attachment; filename=attachment.zip\r\nContent-Transfer-Encoding: base64\r\n\r\n"
                + Convert.ToBase64String(sourceBytes, Base64FormattingOptions.InsertLineBreaks) + "\r\n--probe--\r\n";
            sourceBytes = System.Text.Encoding.UTF8.GetBytes(mime); name = "mail.eml";
        }
        var reader = new OfficeDocumentReaderBuilder().AddImageHandler().AddZipHandler().AddEmailHandler().Build();
        var source = reader.ReadDocument(new MemoryStream(sourceBytes), name);
        string expectedPath = emailContainer ? "mail.eml!/attachment.zip::scan.png" : "attachment.zip::scan.png";
        var engine = new DelegateOcrEngine("tree", (request, _) => {
            Assert.Equal(image, request.Payload);
            Assert.Equal(expectedPath, request.SourceName);
            return Task.FromResult(new OcrResult { Text = "Recognized attachment 42" });
        });
        var result = await source.ApplyOcrTreeAsync(engine);
        Assert.Equal(emailContainer ? 3 : 2, result.Report.DocumentCount);
        Assert.Equal(expectedPath, Assert.Single(result.Recognitions).DocumentPath);
        var content = Assert.Single(result.Document.EnumerateContent(), item => item.Block?.Text == "Recognized attachment 42" || item.Chunk?.Text == "Recognized attachment 42");
        Assert.Equal(expectedPath, content.Location!.Path);
    }

    [Fact]
    public async Task TotalDeadlineStopsAContainerWithoutStartingLaterAttachments() {
        var source = CreateDocument(0);
        source.NestedDocuments = Enumerable.Range(0, 2).Select(index => new OfficeDocumentNestedResult {
            Path = "scan-" + index, Document = CreateDocument(1)
        }).ToArray();
        int calls = 0;
        var engine = new DelegateOcrEngine("deadline", async (_, token) => {
            Interlocked.Increment(ref calls);
            await Task.Delay(Timeout.Infinite, token);
            return new OcrResult();
        });
        var result = await source.ApplyOcrTreeAsync(engine, new OfficeDocumentOcrExecutionOptions {
            TotalTimeout = TimeSpan.FromMilliseconds(100), CandidateTimeout = TimeSpan.FromSeconds(5)
        });
        Assert.InRange(calls, 0, 1);
        Assert.Equal(0, result.Report.RecognizedCandidateCount);
        Assert.Contains(result.Diagnostics, item => item.Code == "ocr-total-time-limit");
        Assert.All(result.Document.NestedDocuments, item => Assert.Single(item.Document.OcrCandidates));
    }

    [Fact]
    public async Task TreeOcrHonorsCallerCancellationAndRejectsCycles() {
        var source = CreateDocument(0);
        source.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "cycle", Document = source } };
        var engine = new DelegateOcrEngine("tree", (_, _) => Task.FromResult(new OcrResult()));
        await Assert.ThrowsAsync<ArgumentException>(() => source.ApplyOcrTreeAsync(engine));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => source.ApplyOcrTreeAsync(engine, cancellationToken: cancellation.Token));
    }
}
