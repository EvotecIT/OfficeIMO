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
    public void KeyValueTable_RendersRichDocumentFactsAndClonesCallerStyle() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 4;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .KeyValueTable(new[] {
                PdfKeyValueRow.Rich(
                    new[] { TextRun.Bolded("Invoice") },
                    new[] { TextRun.Normal("FV/2026/001"), TextRun.Bolded(" paid") }),
                PdfKeyValueRow.Text("Customer", "Evotec")
            }, style: style, includeHeader: true, keyHeader: "Field", valueHeader: "Value")
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Equal(4, style.HeaderRowCount);
        Assert.Contains("Field", text, StringComparison.Ordinal);
        Assert.Contains("Value", text, StringComparison.Ordinal);
        Assert.Contains("Invoice", text, StringComparison.Ordinal);
        Assert.Contains("FV/2026/001 paid", text, StringComparison.Ordinal);
        Assert.Contains("Customer", text, StringComparison.Ordinal);
        Assert.Contains("Evotec", text, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica-Bold", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PageSizes_ResolveExpandedStandardNames() {
        Assert.True(PageSizes.TryGet("a0", out PageSize a0));
        Assert.Equal(2384, a0.Width);
        Assert.Equal(3370, a0.Height);

        Assert.True(PageSizes.TryGet("B10", out PageSize b10));
        Assert.Equal(88, b10.Width);
        Assert.Equal(124, b10.Height);

        Assert.Equal(PageSizes.Executive, PageSizes.Get("Executive"));
        Assert.Equal(PageSizes.Ledger, PageSizes.Get("LedgerOrTabloid"));
        Assert.Contains("Tabloid", PageSizes.Names);
        Assert.False(PageSizes.TryGet("Unknown", out _));
    }

    [Fact]
    public void FileVersion_CanEmitPdf20Header() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf20
            })
            .Paragraph(paragraph => paragraph.Text("PDF 2.0"))
            .ToBytes();

        Assert.StartsWith("%PDF-2.0", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
        Assert.Equal("2.0", PdfDocument.Open(bytes).Inspect().HeaderVersion);
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
    public void DiagnosticsOptimizationTextBlocksFlatteningAndRedactionPlanning_AreAvailableFromDocumentFacade() {
        byte[] source = PdfDocument.Create(new PdfOptions {
                IncludePageLabels = true,
                OpenAction = new PdfOpenActionOptions(1, destinationMode: PdfOpenActionDestinationMode.Fit),
                ViewerPreferences = new PdfViewerPreferencesOptions {
                    DisplayDocTitle = true,
                    FitWindow = true
                }
            })
            .Meta(title: "Diagnostic Report")
            .H1("Diagnostic Report")
            .Paragraph(paragraph => paragraph.Text("This paragraph should appear in structured text and redaction planning."))
            .ToBytes();

        using PdfDocument document = PdfDocument.Open(source);
        PdfDocumentInfo info = document.Inspect();
        Assert.True(info.HasReadablePageLabels);
        Assert.True(info.HasReadableOpenAction);
        Assert.True(info.HasReadableViewerPreferences);

        PdfDiagnosticReport diagnostics = document.Diagnostics();
        Assert.True(diagnostics.CanRead);
        Assert.True(diagnostics.ObjectGraphParsed);
        Assert.True(diagnostics.ObjectCount > 0);
        Assert.True(diagnostics.StreamCount > 0);
        Assert.True(diagnostics.StreamTypeCounts.ContainsKey("Stream"));

        PdfOptimizationReport optimization = document.AnalyzeOptimization();
        Assert.Equal(diagnostics.StreamCount, optimization.StreamCount);
        Assert.NotEmpty(optimization.LargestStreams);

        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Read.TextBlocks();
        Assert.NotEmpty(blocks);
        Assert.Contains(blocks, block => block.Text.Contains("Diagnostic Report", StringComparison.Ordinal));
        Assert.True(document.Read.TryTextBlocks().Succeeded);

        PdfRedactionPlan plan = document.PlanRedactions(new[] {
            new PdfRedactionArea(1, 0, 0, 1000, 1000, "full-page")
        });
        Assert.True(plan.HasMatches);
        Assert.Contains(plan.Matches, match => match.Text != null && match.Text.Contains("structured text", StringComparison.Ordinal));

        PdfDocument flattened = document.FlattenVisualAnnotations();
        Assert.True(flattened.Preflight().CanRead);
        Assert.True(document.TryFlattenVisualAnnotations().Succeeded);
    }

    [Fact]
    public void Forms_TryFillNullOptionsCallsRemainSourceCompatible() {
        byte[] formPdf = BuildSimpleFormPdf();
        var textValues = new Dictionary<string, string> {
            ["Person.Name"] = "Ada Lovelace"
        };
        var fieldValues = new Dictionary<string, PdfFormFieldValue> {
            ["Person.Name"] = PdfFormFieldValue.From("Ada Lovelace")
        };

        PdfDocumentForms forms = PdfDocument.Open(formPdf).Forms;

        Assert.True(forms.TryFill(textValues, null).Succeeded);
        Assert.True(forms.TryFill(fieldValues, null).Succeeded);
        Assert.True(forms.TryFillAndFlatten(textValues, null).Succeeded);
        Assert.True(forms.TryFillAndFlatten(fieldValues, null).Succeeded);
        Assert.True(forms.TryFlatten(null).Succeeded);
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
    public void PageOperations_SplitByPageCountSelectionsAndBookmarks() {
        byte[] source = BuildThreePagePdf();

        IReadOnlyList<PdfDocument> pageGroups = PdfDocument.Open(source).Pages.Split(2);
        Assert.Equal(2, pageGroups.Count);
        Assert.Equal(2, pageGroups[0].Inspect().PageCount);
        Assert.Equal(1, pageGroups[1].Inspect().PageCount);
        Assert.Contains("Page A", pageGroups[0].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Page C", pageGroups[1].Read.Text(), StringComparison.Ordinal);

        IReadOnlyList<PdfDocument> selections = PdfDocument.Open(source).Pages.Split(
            PdfPageSelection.Parse("1-2"),
            PdfPageSelection.Parse("3"));
        Assert.Equal(2, selections.Count);
        Assert.Contains("Second page body", selections[0].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Third page body", selections[1].Read.Text(), StringComparison.Ordinal);

        byte[] bookmarked = BuildThreeBookmarkPdf();
        PdfDocument bookmarkDocument = PdfDocument.Open(bookmarked);
        IReadOnlyList<PdfBookmarkPageRange> ranges = bookmarkDocument.Pages.BookmarkPageRanges();
        Assert.Equal(3, ranges.Count);
        Assert.Equal("Chapter One", ranges[0].Title);
        Assert.Equal(PdfPageRange.From(1, 1), ranges[0].PageRange);
        Assert.Equal(PdfPageRange.From(3, 3), ranges[2].PageRange);

        IReadOnlyList<PdfBookmarkPageRange> selectedRange = bookmarkDocument.Pages.BookmarkPageRanges("Chapter Two");
        Assert.Equal(PdfPageRange.From(2, 2), Assert.Single(selectedRange).PageRange);

        IReadOnlyList<PdfDocument> bookmarkSplit = bookmarkDocument.Pages.SplitByBookmarks("Chapter Two");
        PdfDocument chapterTwo = Assert.Single(bookmarkSplit);
        Assert.Contains("Chapter Two", chapterTwo.Read.Text(), StringComparison.Ordinal);
        Assert.DoesNotContain("Chapter Three", chapterTwo.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void PageOperations_BookmarkRangesFollowPageOrderWhenOutlineTreeIsOutOfOrder() {
        PdfDocument bookmarkDocument = PdfDocument.Open(BuildOutOfOrderBookmarkPdf());

        IReadOnlyList<PdfBookmarkPageRange> ranges = bookmarkDocument.Pages.BookmarkPageRanges();

        Assert.Equal(new[] { "Chapter One", "Chapter Two", "Chapter Three" }, ranges.Select(range => range.Title).ToArray());
        Assert.Equal(PdfPageRange.From(1, 1), ranges[0].PageRange);
        Assert.Equal(PdfPageRange.From(2, 2), ranges[1].PageRange);
        Assert.Equal(PdfPageRange.From(3, 3), ranges[2].PageRange);
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

        Assert.Empty(document.AnalyzeTextEncoding());

        PdfBytesResult bytesResult = document.TryToBytes();

        Assert.True(bytesResult.Succeeded);
        Assert.Equal(invalidPdf.LongLength, bytesResult.ByteCount);
        Assert.Equal(invalidPdf, bytesResult.Bytes);
        Assert.Equal(invalidPdf, bytesResult.RequireBytes());
        Assert.Empty(bytesResult.Diagnostics);
        Assert.Empty(bytesResult.TextEncodingDiagnostics);
        Assert.Empty(bytesResult.ConversionWarnings);

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

    [Fact]
    public void SaveResult_CarriesTextEncodingDiagnosticsForGeneratedPdfFailures() {
        using PdfDocument document = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Snowman \u2603"));
        using var stream = new MemoryStream();

        PdfBytesResult bytesResult = document.TryToBytes();
        Assert.False(bytesResult.Succeeded);
        Assert.Equal(0, bytesResult.ByteCount);
        Assert.Empty(bytesResult.Bytes);
        Assert.NotEmpty(bytesResult.Diagnostics);
        Assert.Throws<InvalidOperationException>(() => bytesResult.RequireBytes());

        PdfSaveResult result = document.TrySave(stream);

        Assert.False(result.Succeeded);
        Assert.Equal(0, result.BytesWritten);
        Assert.Equal(0, stream.Length);
        Assert.NotEmpty(result.Diagnostics);

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(result.TextEncodingDiagnostics);
        PdfTextEncodingDiagnostic bytesDiagnostic = Assert.Single(bytesResult.TextEncodingDiagnostics);

        Assert.Equal("unsupported-text-glyph", diagnostic.Code);
        Assert.Equal("PdfParagraph", diagnostic.Source);
        Assert.Equal("PdfParagraph[0].Run[0]", diagnostic.Location);
        Assert.Equal(0, diagnostic.RunIndex);
        Assert.Equal("U+2603", diagnostic.CodePoint);
        Assert.Equal("\u2603", diagnostic.Text);
        Assert.False(diagnostic.IsControlCharacter);
        Assert.Equal("PDF WinAnsiEncoding", diagnostic.Encoding);
        Assert.Equal("Embedded Unicode fonts are required for this text.", diagnostic.Remediation);
        Assert.Equal(diagnostic.CodePoint, bytesDiagnostic.CodePoint);
        Assert.Equal(diagnostic.Location, bytesDiagnostic.Location);

        PdfConversionWarning warning = Assert.Single(result.ConversionWarnings);
        PdfConversionWarning bytesWarning = Assert.Single(bytesResult.ConversionWarnings);

        Assert.Equal("OfficeIMO.Pdf", warning.Converter);
        Assert.Equal(diagnostic.Code, warning.Code);
        Assert.Equal(diagnostic.Message, warning.Message);
        Assert.Equal(PdfConversionWarningSeverity.Error, warning.Severity);
        Assert.Equal("U+2603", warning.Details["codePoint"]);
        Assert.Equal("PdfParagraph[0].Run[0]", warning.Details["location"]);
        Assert.Equal("0", warning.Details["runIndex"]);
        Assert.Equal("PDF WinAnsiEncoding", warning.Details["encoding"]);
        Assert.Equal("Embedded Unicode fonts are required for this text.", warning.Details["remediation"]);
        Assert.Equal(warning.Code, bytesWarning.Code);
    }

    [Fact]
    public void AnalyzeTextEncoding_ReturnsAllGeneratedTextDiagnosticsBeforeRender() {
        var options = new PdfOptions {
            ShowHeader = true,
            HeaderFormat = "Header \u2603",
            ShowPageNumbers = true,
            FooterFormat = "Footer \u2602"
        };

        using PdfDocument document = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Paragraph \u2603"))
            .H1("Heading \U0001F680")
            .Bullets(new[] { "Bullet \u266b" })
            .Table(new[] { new[] { "Table \u25a0" } }, style: new PdfTableStyle {
                Caption = "Table caption \u2666"
            })
            .Canvas(canvas => canvas
                .Text("Canvas \u260e", 10, 10, 120, 24)
                .FreeTextAnnotation("Callout \u2615", 10, 42, 120, 32)
                .Table(new[] { new[] { "Canvas table" } }, 10, 84, 120, 42, new PdfTableStyle {
                    Caption = "Canvas caption \u273f"
                }))
            .TextField("Person.Name", width: 120, height: 20, value: "Field \u2603");

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        Assert.Equal(10, diagnostics.Count);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfHeader" && diagnostic.Location == "PdfHeader[page=1]" && diagnostic.PageNumber == 1 && diagnostic.CodePoint == "U+2603");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfFooter" && diagnostic.Location == "PdfFooter[page=1]" && diagnostic.PageNumber == 1 && diagnostic.CodePoint == "U+2602");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfParagraph" && diagnostic.Location == "PdfParagraph[0].Run[0]" && diagnostic.RunIndex == 0 && diagnostic.CodePoint == "U+2603");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfHeading" && diagnostic.Location == "PdfHeading[1]" && diagnostic.CodePoint == "U+1F680");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfListItem" && diagnostic.Location == "PdfBulletList[2].PdfListItem[0].Run[0]" && diagnostic.RunIndex == 0 && diagnostic.CodePoint == "U+266B");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfTableCaption" && diagnostic.Location == "PdfTable[3].PdfTableCaption" && diagnostic.CodePoint == "U+2666");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfTableCell" && diagnostic.Location == "PdfTable[3].PdfTableCell[0,0].Run[0]" && diagnostic.RunIndex == 0 && diagnostic.TableRowIndex == 0 && diagnostic.TableColumnIndex == 0 && diagnostic.CodePoint == "U+25A0");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfCanvasText" && diagnostic.Location == "PdfCanvas[4].PdfCanvasText[0].Run[0]" && diagnostic.RunIndex == 0 && diagnostic.CodePoint == "U+260E");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfTableCaption" && diagnostic.Location == "PdfCanvas[4].PdfCanvasTable[2].PdfTableCaption" && diagnostic.CodePoint == "U+273F");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfTextField" && diagnostic.Location == "PdfTextField[5]" && diagnostic.FieldName == "Person.Name" && diagnostic.CodePoint == "U+2603");

        PdfBytesResult result = document.TryToBytes();
        using var stream = new MemoryStream();
        PdfSaveResult saveResult = document.TrySave(stream);

        Assert.False(result.Succeeded);
        Assert.Equal(diagnostics.Count, result.TextEncodingDiagnostics.Count);
        Assert.Equal("PdfHeader", result.TextEncodingDiagnostics[0].Source);
        Assert.Equal("PdfHeader[page=1]", result.TextEncodingDiagnostics[0].Location);
        Assert.Equal(1, result.TextEncodingDiagnostics[0].PageNumber);
        Assert.Equal("U+2603", result.TextEncodingDiagnostics[0].CodePoint);
        Assert.Equal("PDF WinAnsiEncoding", result.TextEncodingDiagnostics[0].Encoding);
        Assert.Equal(diagnostics.Count, result.ConversionWarnings.Count);
        Assert.Equal("PdfHeader[page=1]", result.ConversionWarnings[0].Details["location"]);
        Assert.Equal("1", result.ConversionWarnings[0].Details["pageNumber"]);
        Assert.Equal("PDF WinAnsiEncoding", result.ConversionWarnings[0].Details["encoding"]);
        Assert.Contains(result.ConversionWarnings, warning =>
            warning.Source == "PdfTableCell" &&
            warning.Details["tableRowIndex"] == "0" &&
            warning.Details["tableColumnIndex"] == "0");
        Assert.Contains(result.ConversionWarnings, warning =>
            warning.Source == "PdfTextField" &&
            warning.Details["fieldName"] == "Person.Name");
        Assert.Contains("preflight found 10 generated text issues", result.Diagnostics[0], StringComparison.Ordinal);

        ArgumentException preflightException = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());

        Assert.Equal(1, preflightException.Data["pageNumber"]);

        Assert.False(saveResult.Succeeded);
        Assert.Equal(diagnostics.Count, saveResult.TextEncodingDiagnostics.Count);
        Assert.Equal(0, saveResult.BytesWritten);
        Assert.Equal(0, stream.Length);
    }

    [Fact]
    public void AnalyzeTextEncoding_PreflightsVariantTextWatermarks() {
        var options = new PdfOptions {
            FirstPageTextWatermark = new PdfTextWatermark("First \u2603"),
            EvenPageTextWatermark = new PdfTextWatermark("Even \u2602")
        };

        using PdfDocument document = PdfDocument.Create(options);
        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfFirstPageTextWatermark" && diagnostic.Location == "PdfFirstPageTextWatermark" && diagnostic.CodePoint == "U+2603");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfEvenPageTextWatermark" && diagnostic.Location == "PdfEvenPageTextWatermark" && diagnostic.CodePoint == "U+2602");
    }

    [Fact]
    public void AnalyzeTextEncoding_PreflightsFormWidgetValuesThroughEmbeddedHelveticaPath() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions();
        options.EmbedStandardFont(PdfStandardFont.Helvetica, fontPath);
        const string value = "\u0105";
        if (PdfTextDiagnostics.AnalyzeGeneratedText(value, options, PdfStandardFont.Helvetica).Count != 0) {
            return;
        }

        using PdfDocument document = PdfDocument.Create(options)
            .TextField("Person.City", value: value)
            .ChoiceField("Person.Country", new[] { "PL", value }, value: value);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.FieldName == "Person.City" && diagnostic.CodePoint == "U+0105");
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.FieldName == "Person.Country" && diagnostic.CodePoint == "U+0105");

        using var stream = new MemoryStream();
        PdfSaveResult result = document.TrySave(stream);

        Assert.True(result.Succeeded);
        Assert.True(stream.Length > 0);
        Assert.DoesNotContain(result.TextEncodingDiagnostics, diagnostic => diagnostic.FieldName == "Person.City" && diagnostic.CodePoint == "U+0105");
    }

    [Fact]
    public void AnalyzeTextEncoding_PreflightsFreeTextAppearanceThroughWinAnsiPath() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions();
        options.EmbedStandardFont(PdfStandardFont.Helvetica, fontPath);
        const string value = "\u0105";
        if (PdfTextDiagnostics.AnalyzeGeneratedText(value, options, PdfStandardFont.Helvetica).Count != 0) {
            return;
        }

        using PdfDocument document = PdfDocument.Create(options)
            .FreeTextAnnotation(value, width: 120, height: 32);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfFreeTextAnnotation" && diagnostic.CodePoint == "U+0105");

        using var stream = new MemoryStream();
        PdfSaveResult result = document.TrySave(stream);

        Assert.False(result.Succeeded);
        Assert.Equal(0, stream.Length);
        Assert.Contains(result.TextEncodingDiagnostics, diagnostic => diagnostic.Source == "PdfFreeTextAnnotation" && diagnostic.CodePoint == "U+0105");
    }

    [Fact]
    public void ToBytes_ThrowsFullGeneratedTextPreflightBeforeRendering() {
        using PdfDocument document = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Paragraph \u2603"))
            .H1("Heading \U0001F680");

        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());

        Assert.Contains("preflight found 2 generated text issues", exception.Message, StringComparison.Ordinal);
        Assert.Contains("U+2603", exception.Message, StringComparison.Ordinal);
        Assert.Equal("unsupported-text-glyph", exception.Data["code"]);
        Assert.Equal("PdfParagraph", exception.Data["source"]);
        Assert.Equal("PdfParagraph[0].Run[0]", exception.Data["location"]);
        Assert.Equal(0, exception.Data["runIndex"]);
        Assert.Equal(2, exception.Data["diagnosticsCount"]);

        var diagnostics = Assert.IsAssignableFrom<IReadOnlyList<PdfTextEncodingDiagnostic>>(exception.Data["textEncodingDiagnostics"]);
        Assert.Equal(2, diagnostics.Count);
        Assert.Equal("PdfParagraph[0].Run[0]", diagnostics[0].Location);
        Assert.Equal("PdfHeading[1]", diagnostics[1].Location);
    }

    [Fact]
    public void ToBytes_ThrowsTableCellCoordinatesInTextPreflightException() {
        using PdfDocument document = PdfDocument.Create()
            .Table(new[] { new[] { "Table \u25a0" } });

        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());

        Assert.Equal("PdfTableCell", exception.Data["source"]);
        Assert.Equal("PdfTable[0].PdfTableCell[0,0].Run[0]", exception.Data["location"]);
        Assert.Equal(0, exception.Data["tableRowIndex"]);
        Assert.Equal(0, exception.Data["tableColumnIndex"]);

        var diagnostics = Assert.IsAssignableFrom<IReadOnlyList<PdfTextEncodingDiagnostic>>(exception.Data["textEncodingDiagnostics"]);
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnostics);

        Assert.Equal(0, diagnostic.TableRowIndex);
        Assert.Equal(0, diagnostic.TableColumnIndex);
    }

    [Fact]
    public void ToBytes_ThrowsFieldNameForGeneratedFormPreflightException() {
        using PdfDocument document = PdfDocument.Create()
            .Table(new[] {
                new[] {
                    PdfTableCell.WithFormFields(
                        "Table form",
                        new[] {
                            PdfTableCellFormField.TextField("Table.DueDate", "Due \u2603", width: 120, height: 18)
                        })
                }
            });

        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());

        Assert.Equal("PdfTableTextField", exception.Data["source"]);
        Assert.Equal("PdfTable[0].PdfTableCell[0,0].PdfTableTextField", exception.Data["location"]);
        Assert.Equal("Table.DueDate", exception.Data["fieldName"]);
        Assert.Equal(0, exception.Data["tableRowIndex"]);
        Assert.Equal(0, exception.Data["tableColumnIndex"]);

        var diagnostics = Assert.IsAssignableFrom<IReadOnlyList<PdfTextEncodingDiagnostic>>(exception.Data["textEncodingDiagnostics"]);
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnostics);

        Assert.Equal("Table.DueDate", diagnostic.FieldName);
        Assert.Equal(0, diagnostic.TableRowIndex);
        Assert.Equal(0, diagnostic.TableColumnIndex);
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

    private static byte[] BuildThreeBookmarkPdf() {
        return PdfDocument.Create(new PdfOptions {
                CreateOutlineFromHeadings = true
            })
            .H1("Chapter One")
            .Paragraph(p => p.Text("First chapter body"))
            .PageBreak()
            .H1("Chapter Two")
            .Paragraph(p => p.Text("Second chapter body"))
            .PageBreak()
            .H1("Chapter Three")
            .Paragraph(p => p.Text("Third chapter body"))
            .ToBytes();
    }

    private static byte[] BuildOutOfOrderBookmarkPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /Outlines 10 0 R /PageMode /UseOutlines >>",
            "<< /Type /Pages /Count 3 /Kids [3 0 R 4 0 R 5 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 9 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 9 0 R >> >> /Contents 7 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 9 0 R >> >> /Contents 8 0 R >>",
            BuildWorkflowStream("BT\n/F1 12 Tf\n72 200 Td\n(Chapter One body) Tj\nET"),
            BuildWorkflowStream("BT\n/F1 12 Tf\n72 200 Td\n(Chapter Two body) Tj\nET"),
            BuildWorkflowStream("BT\n/F1 12 Tf\n72 200 Td\n(Chapter Three body) Tj\nET"),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "<< /Type /Outlines /First 11 0 R /Last 13 0 R /Count 3 >>",
            "<< /Title (Chapter Three) /Parent 10 0 R /Next 12 0 R /Dest [5 0 R /XYZ 0 200 0] >>",
            "<< /Title (Chapter One) /Parent 10 0 R /Prev 11 0 R /Next 13 0 R /Dest [3 0 R /XYZ 0 200 0] >>",
            "<< /Title (Chapter Two) /Parent 10 0 R /Prev 12 0 R /Dest [4 0 R /XYZ 0 200 0] >>"
        };

        return Encoding.ASCII.GetBytes(BuildWorkflowPdf(objects));
    }

    private static string BuildWorkflowStream(string content) {
        byte[] bytes = Encoding.ASCII.GetBytes(content);
        return "<< /Length " + bytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream";
    }

    private static string BuildWorkflowPdf(IReadOnlyList<string> objects) {
        var builder = new StringBuilder();
        builder.AppendLine("%PDF-1.7");
        for (int i = 0; i < objects.Count; i++) {
            builder.Append((i + 1).ToString(CultureInfo.InvariantCulture)).AppendLine(" 0 obj");
            builder.AppendLine(objects[i]);
            builder.AppendLine("endobj");
        }

        builder.AppendLine("trailer");
        builder.Append("<< /Root 1 0 R /Size ").Append(objects.Count + 1).AppendLine(" >>");
        builder.AppendLine("startxref");
        builder.AppendLine("123");
        builder.AppendLine("%%EOF");
        return builder.ToString();
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
