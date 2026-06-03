using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentWorkflowTests {
    [Fact]
    public void PageSelection_ParsesAndSnapshotsCallerRanges() {
        PdfPageSelection parsed = PdfPageSelection.Parse("3,1-2;2..3");

        Assert.Equal(5, parsed.PageCount);
        Assert.Equal("3,1-2,2-3", parsed.ToString());
        Assert.Equal(new[] {
            PdfPageRange.From(3, 3),
            PdfPageRange.From(1, 2),
            PdfPageRange.From(2, 3)
        }, parsed.Ranges);

        var ranges = new[] { PdfPageRange.From(1, 1) };
        PdfPageSelection selection = PdfPageSelection.FromRanges(ranges);
        ranges[0] = PdfPageRange.From(2, 2);

        Assert.Equal("1", selection.ToString());
        Assert.True(PdfPageSelection.TryParse("1,3", out PdfPageSelection? tryParsed));
        Assert.Equal(PdfPageSelection.FromRanges(PdfPageRange.From(1, 1), PdfPageRange.From(3, 3)), tryParsed);
        Assert.False(PdfPageSelection.TryParse(" ", out _));
    }

    [Fact]
    public void Open_SnapshotsInputBytesAndExposesReadInspectAndPreflight() {
        byte[] source = BuildThreePagePdf();
        byte[] callerBuffer = (byte[])source.Clone();

        using PdfDocument document = PdfDocument.Open(callerBuffer);
        callerBuffer[20] ^= 0x10;

        Assert.Equal(3, document.Inspect().PageCount);
        Assert.Equal("Workflow source", document.Inspect().Metadata.Title);
        Assert.Equal(PdfTextExtractor.ExtractAllText(source), document.Read.Text());
        Assert.Equal(PdfTextExtractor.ExtractTextByPage(source), document.Read.TextByPage());
        Assert.True(document.Preflight().CanRead);
        Assert.True(document.Preflight().CanRewrite);
    }

    [Fact]
    public void PageSelection_DrivesPageAndReadWorkflows() {
        byte[] source = BuildThreePagePdf();
        PdfPageSelection selection = PdfPageSelection.Parse("3,1-2");

        PdfDocument extracted = PdfDocument.Open(source).Pages.Extract(selection);
        Assert.Equal(PdfDocument.Open(source).Pages.Extract("3,1-2").ToBytes(), extracted.ToBytes());
        Assert.Equal(3, extracted.Inspect().PageCount);
        Assert.Contains("Page C", extracted.Read.Text(), StringComparison.Ordinal);

        Assert.Equal(
            PdfDocument.Open(source).Pages.Delete("2").ToBytes(),
            PdfDocument.Open(source).Pages.Delete(PdfPageSelection.From(2)).ToBytes());

        Assert.Equal(
            PdfDocument.Open(source).Pages.Reorder("2,3,1").ToBytes(),
            PdfDocument.Open(source).Pages.Reorder(PdfPageSelection.Parse("2,3,1")).ToBytes());

        Assert.Equal(
            PdfDocument.Open(source).Pages.Duplicate("2").ToBytes(),
            PdfDocument.Open(source).Pages.Duplicate(PdfPageSelection.From(PdfPageRange.From(2, 2))).ToBytes());

        Assert.Equal(
            PdfDocument.Open(source).Pages.Move(1, "3").ToBytes(),
            PdfDocument.Open(source).Pages.Move(1, PdfPageSelection.From(3)).ToBytes());

        Assert.Equal(
            PdfDocument.Open(source).Pages.Rotate(90, "2").ToBytes(),
            PdfDocument.Open(source).Pages.Rotate(90, PdfPageSelection.Parse("2")).ToBytes());

        PdfDocument opened = PdfDocument.Open(source);
        Assert.Equal(PdfTextExtractor.ExtractAllTextByPageRanges(source, PdfPageRange.ParseMany("2,1")), opened.Read.Text(PdfPageSelection.Parse("2,1")));
        Assert.Equal(PdfTextExtractor.ExtractTextByPageRanges(source, PdfPageRange.ParseMany("2,1")), opened.Read.TextByPage(PdfPageSelection.Parse("2,1")));
        Assert.Equal(2, opened.Read.Logical(PdfPageSelection.Parse("2,1")).Pages.Count);
        Assert.Contains("Second page body", opened.Read.Markdown(PdfPageSelection.Parse("2")), StringComparison.Ordinal);
    }

    [Fact]
    public void OperationResult_PreflightsPageOperationsAndCarriesDiagnostics() {
        byte[] source = BuildThreePagePdf();

        PdfOperationResult<PdfDocument> extracted = PdfDocument.Open(source).Pages.TryExtract(PdfPageSelection.Parse("2"));
        Assert.True(extracted.CanAttempt);
        Assert.True(extracted.Succeeded);
        Assert.Empty(extracted.Diagnostics);
        Assert.Equal(PdfPreflightCapability.ManipulatePages, extracted.Capability);
        Assert.Contains("Page B", extracted.RequireValue().Read.Text(), StringComparison.Ordinal);

        PdfOperationResult<IReadOnlyList<PdfDocument>> split = PdfDocument.Open(source).Pages.TrySplit();
        Assert.True(split.Succeeded);
        Assert.Equal(3, split.RequireValue().Count);

        PdfDocument invalid = PdfDocument.Open(Encoding.ASCII.GetBytes("not a pdf"));
        PdfOperationResult<PdfDocument> blocked = invalid.Pages.TryExtract(PdfPageSelection.From(1));

        Assert.False(blocked.CanAttempt);
        Assert.False(blocked.Succeeded);
        Assert.Null(blocked.Value);
        Assert.NotEmpty(blocked.Diagnostics);
        Assert.NotNull(blocked.Preflight);
        Assert.Throws<InvalidOperationException>(() => blocked.RequireValue());
    }

    [Fact]
    public void OperationResult_ExtendsAcrossMergeReadStampAndForms() {
        byte[] source = BuildThreePagePdf();
        byte[] appendix = BuildPdf("Appendix", "Appendix body");

        PdfOperationResult<PdfDocument> merged = PdfDocument.Open(source).TryMergeWith(PdfDocument.Open(appendix));
        Assert.True(merged.Succeeded);
        Assert.Equal(4, merged.RequireValue().Inspect().PageCount);

        PdfDocument opened = PdfDocument.Open(source);
        PdfOperationResult<string> text = opened.Read.TryText(PdfPageSelection.Parse("2"));
        Assert.True(text.Succeeded);
        Assert.Contains("Second page body", text.RequireValue(), StringComparison.Ordinal);

        PdfOperationResult<PdfLogicalDocument> logical = opened.Read.TryLogical(PdfPageSelection.Parse("1,3"));
        Assert.True(logical.Succeeded);
        Assert.Equal(2, logical.RequireValue().Pages.Count);

        PdfOperationResult<string> markdown = opened.Read.TryMarkdown(PdfPageSelection.Parse("1"));
        Assert.True(markdown.Succeeded);
        Assert.Contains("First page body", markdown.RequireValue(), StringComparison.Ordinal);

        PdfOperationResult<PdfDocument> stamped = opened.Stamp.TryText("Reviewed", new PdfTextStampOptions { X = 72, Y = 72 });
        Assert.True(stamped.Succeeded);
        Assert.Equal(3, stamped.RequireValue().Inspect().PageCount);

        byte[] formPdf = BuildSimpleFormPdf();
        PdfOperationResult<PdfDocument> filled = PdfDocument.Open(formPdf).Forms.TryFill(new Dictionary<string, string> {
            ["Person.Name"] = "Ada Lovelace"
        });
        Assert.True(filled.Succeeded);
        Assert.Equal("Ada Lovelace", Assert.Single(filled.RequireValue().Inspect().FormFields).Value);

        PdfOperationResult<PdfDocument> flattened = PdfDocument.Open(formPdf).Forms.TryFillAndFlatten(new Dictionary<string, string> {
            ["Person.Name"] = "Ada Lovelace"
        });
        Assert.True(flattened.Succeeded);
        Assert.Empty(flattened.RequireValue().Inspect().FormFields);

        PdfDocument invalid = PdfDocument.Open(Encoding.ASCII.GetBytes("not a pdf"));
        PdfOperationResult<string> blockedText = invalid.Read.TryText();
        Assert.False(blockedText.CanAttempt);
        Assert.NotEmpty(blockedText.Diagnostics);

        PdfOperationResult<PdfDocument> blockedStamp = invalid.Stamp.TryText("Reviewed");
        Assert.False(blockedStamp.CanAttempt);
        Assert.NotEmpty(blockedStamp.Diagnostics);
    }

    [Fact]
    public void PageOperations_ReturnNewDocumentsAndMatchExistingHelpers() {
        byte[] source = BuildThreePagePdf();

        Assert.Equal(
            PdfPageExtractor.ExtractPageRanges(source, PdfPageRange.ParseMany("3,1-2")),
            PdfDocument.Open(source).Pages.Extract("3,1-2").ToBytes());

        Assert.Equal(
            PdfPageEditor.DeletePageRanges(source, PdfPageRange.ParseMany("2")),
            PdfDocument.Open(source).Pages.Delete("2").ToBytes());

        Assert.Equal(
            PdfPageEditor.ReorderPageRanges(source, PdfPageRange.ParseMany("2,3,1")),
            PdfDocument.Open(source).Pages.Reorder("2,3,1").ToBytes());

        Assert.Equal(
            PdfPageEditor.RotatePageRanges(source, 90, PdfPageRange.ParseMany("2")),
            PdfDocument.Open(source).Pages.Rotate(90, "2").ToBytes());

        IReadOnlyList<PdfDocument> split = PdfDocument.Open(source).Pages.Split();
        Assert.Equal(3, split.Count);
        Assert.All(split, part => Assert.Equal(1, part.Inspect().PageCount));
        Assert.Contains("Page A", split[0].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Page B", split[1].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Page C", split[2].Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void MergeMetadataAndStamping_StayFluentAndDelegateToCurrentEngine() {
        byte[] source = BuildThreePagePdf();
        byte[] appendix = BuildPdf("Appendix", "Appendix body");

        PdfDocument merged = PdfDocument.Open(source).MergeWith(PdfDocument.Open(appendix));
        Assert.Equal(PdfMerger.Merge(source, appendix), merged.ToBytes());
        Assert.Equal(4, merged.Inspect().PageCount);

        PdfDocument metadata = merged.UpdateMetadata(title: "Workflow updated", author: "OfficeIMO Tests");
        Assert.Equal(
            PdfMetadataEditor.UpdateMetadata(merged.ToBytes(), title: "Workflow updated", author: "OfficeIMO Tests"),
            metadata.ToBytes());
        Assert.Equal("Workflow updated", metadata.Inspect().Metadata.Title);
        Assert.Equal("OfficeIMO Tests", metadata.Inspect().Metadata.Author);

        var stampOptions = new PdfTextStampOptions {
            X = 72,
            Y = 72,
            FontSize = 12
        };

        Assert.Equal(
            PdfStamper.StampText(metadata.ToBytes(), "Reviewed", stampOptions),
            metadata.Stamp.Text("Reviewed", stampOptions).ToBytes());
    }

    [Fact]
    public void Save_WritesCurrentBytesToStreamAndPath() {
        using PdfDocument document = PdfDocument.Open(BuildThreePagePdf()).Pages.Delete(2);
        using var stream = new MemoryStream();

        PdfDocument returned = document.Save(stream);

        Assert.Same(document, returned);
        Assert.Equal(document.ToBytes(), stream.ToArray());

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-workflow-" + Guid.NewGuid().ToString("N"));
        string path = Path.Combine(directory, "saved.pdf");
        try {
            document.Save(path);

            Assert.True(File.Exists(path));
            Assert.Equal(document.ToBytes(), File.ReadAllBytes(path));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public async System.Threading.Tasks.Task SaveResult_ReportsOutputWithoutRequiringReadablePdfContent() {
        byte[] invalidPdf = Encoding.ASCII.GetBytes("not a pdf");
        using PdfDocument document = PdfDocument.Open(invalidPdf);
        using var stream = new MemoryStream();

        PdfSaveResult streamResult = document.TrySave(stream);

        Assert.True(streamResult.Succeeded);
        Assert.Null(streamResult.OutputPath);
        Assert.Equal(invalidPdf.LongLength, streamResult.BytesWritten);
        Assert.Empty(streamResult.Diagnostics);
        Assert.Same(streamResult, streamResult.RequireSuccess());
        Assert.Equal(invalidPdf, stream.ToArray());

        using var asyncStream = new MemoryStream();
        PdfSaveResult asyncResult = await document.TrySaveAsync(asyncStream);

        Assert.True(asyncResult.Succeeded);
        Assert.Equal(invalidPdf.LongLength, asyncResult.BytesWritten);
        Assert.Equal(invalidPdf, asyncStream.ToArray());

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-save-result-" + Guid.NewGuid().ToString("N"));
        string path = Path.Combine(directory, "snapshot.pdf");
        try {
            PdfSaveResult pathResult = document.TrySave(path);

            Assert.True(pathResult.Succeeded);
            Assert.Equal(Path.GetFullPath(path), pathResult.OutputPath);
            Assert.Equal(invalidPdf.LongLength, pathResult.BytesWritten);
            Assert.Equal(invalidPdf, File.ReadAllBytes(path));

            PdfSaveResult directoryResult = document.TrySave(directory);

            Assert.False(directoryResult.Succeeded);
            Assert.Equal(0, directoryResult.BytesWritten);
            Assert.NotEmpty(directoryResult.Diagnostics);
            Assert.Throws<InvalidOperationException>(() => directoryResult.RequireSuccess());
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }

        using var readOnlyStream = new MemoryStream(Array.Empty<byte>(), writable: false);
        PdfSaveResult streamFailure = document.TrySave(readOnlyStream);

        Assert.False(streamFailure.Succeeded);
        Assert.NotEmpty(streamFailure.Diagnostics);
    }

    private static byte[] BuildThreePagePdf() {
        return PdfDocument.Create()
            .Meta(title: "Workflow source", author: "OfficeIMO")
            .H1("Page A")
            .Paragraph(p => p.Text("First page body"))
            .PageBreak()
            .H1("Page B")
            .Paragraph(p => p.Text("Second page body"))
            .PageBreak()
            .H1("Page C")
            .Paragraph(p => p.Text("Third page body"))
            .ToBytes();
    }

    private static byte[] BuildPdf(string title, string text) {
        return PdfDocument.Create()
            .Meta(title: title, author: "OfficeIMO")
            .H1(title)
            .Paragraph(p => p.Text(text))
            .ToBytes();
    }

    private static byte[] BuildSimpleFormPdf() {
        return PdfDocument.Create()
            .H1("Form")
            .TextField("Person.Name", width: 180, height: 24, value: "Original")
            .ToBytes();
    }
}
