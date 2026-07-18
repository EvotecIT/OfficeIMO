using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageImporterTests {
    [Fact]
    public void AppendPages_ImportsSelectedSourcePagesInRequestedOrder() {
        byte[] target = BuildPdf("Target", "Target page");
        byte[] source = BuildPdf("Source", "Source first", "Source second", "Source third");

        byte[] imported = PdfPageImporter.AppendPages(target, source, 3, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(3, info.PageCount);
        Assert.Equal("Target", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetpage", "Sourcethird", "Sourcefirst");
        Assert.DoesNotContain("Sourcesecond", text);
    }

    [Fact]
    public void AppendPages_ImportsDuplicateSourceSelectionsAsClonedPages() {
        byte[] target = BuildPdf("Target", "Target page");
        byte[] source = BuildPdf("Source", "Source first", "Source second");

        byte[] imported = PdfPageImporter.AppendPages(target, source, 2, 2, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(4, info.PageCount);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetpage", "Sourcesecond", "Sourcesecond", "Sourcefirst");
        Assert.Equal(2, CountOccurrences(text, "Sourcesecond"));
    }

    [Fact]
    public void AppendPages_WithFlattenVisualAnnotationsOption_FlattensImportedSourceMarkupOnly() {
        byte[] target = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Meta(title: "Target", author: "OfficeIMO")
            .Paragraph(p => p.Text("Target page"))
            .TextAnnotation("Target note")
            .ToBytes();
        byte[] source = BuildAnnotatedPdf("Source", "Annotated source page");

        byte[] imported = PdfPageImporter.AppendPages(
            new PdfPageImportOptions {
                FlattenVisualAnnotations = true
            },
            target,
            source);
        string pdf = Encoding.ASCII.GetString(imported);
        PdfDocumentInfo info = PdfInspector.Inspect(imported);

        Assert.Equal(2, info.PageCount);
        Assert.Equal(1, info.AnnotationCount);
        Assert.Single(info.GetAnnotationsBySubtype("Text"));
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Highlight", pdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Text", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetpage", "Annotatedsourcepage");
    }

    [Fact]
    public void PrependPages_ImportsAllSourcePagesWhenSelectionIsEmpty() {
        byte[] target = BuildPdf("Target", "Target first", "Target second");
        byte[] source = BuildPdf("Source", "Source first", "Source second");

        byte[] imported = PdfPageImporter.PrependPages(target, source);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(4, info.PageCount);
        Assert.Equal("Target", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Sourcefirst", "Sourcesecond", "Targetfirst", "Targetsecond");
    }

    [Fact]
    public void InsertPages_RebasesTargetPageLabelsAfterInsertedPagesAtBeginning() {
        byte[] target = BuildTwoPageLabelPdf();
        byte[] source = BuildPdf("Inserted", "Inserted page");

        byte[] imported = PdfPageImporter.InsertPages(target, source, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(3, info.PageCount);
        PdfPageLabel label = Assert.Single(info.PageLabels);
        Assert.Equal(1, label.StartPageIndex);
        Assert.Equal(2, label.StartPageNumber);
        Assert.Equal("D", label.Style);
        Assert.Equal(1, label.StartNumber);

        string text = Encoding.ASCII.GetString(imported);
        Assert.Contains("/PageLabels ", text, StringComparison.Ordinal);
        Assert.Contains("/Nums [ 1 << /S /D /St 1 >> ]", text, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertPages_RebasesTargetPageLabelsAroundInsertedMiddlePages() {
        byte[] target = BuildTwoPageLabelPdf();
        byte[] source = BuildPdf("Inserted", "Inserted page");

        byte[] imported = PdfPageImporter.InsertPages(target, source, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(3, info.PageCount);

        string text = Encoding.ASCII.GetString(imported);
        Assert.Contains("/Nums [ 0 << /S /D /St 1 >> 2 << /S /D /St 2 >> ]", text, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertPages_ImportsSelectedSourcePagesBeforeTargetPage() {
        byte[] target = BuildPdf("Target", "Target first", "Target second", "Target third");
        byte[] source = BuildPdf("Source", "Source first", "Source second", "Source third");

        byte[] imported = PdfPageImporter.InsertPages(target, source, 2, 3, 1);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(5, info.PageCount);
        Assert.Equal("Target", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetfirst", "Sourcethird", "Sourcefirst", "Targetsecond", "Targetthird");
        Assert.DoesNotContain("Sourcesecond", text);
    }

    [Fact]
    public void InsertPages_UsesPageCountPlusOneToInsertAtEnd() {
        byte[] target = BuildPdf("Target", "Target first", "Target second");
        byte[] source = BuildPdf("Source", "Source first", "Source second");

        byte[] imported = PdfPageImporter.InsertPages(target, source, 3, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(3, info.PageCount);
        Assert.Equal("Target", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetfirst", "Targetsecond", "Sourcesecond");
        Assert.DoesNotContain("Sourcefirst", text);
    }

    [Fact]
    public void InsertPages_PreservesTargetMetadataWhenInsertingAtBeginning() {
        byte[] target = BuildPdf("Target", "Target first", "Target second");
        byte[] source = BuildPdf("Source", "Source first", "Source second");

        byte[] imported = PdfPageImporter.InsertPages(target, source, 1, 2);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(3, info.PageCount);
        Assert.Equal("Target", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Sourcesecond", "Targetfirst", "Targetsecond");
        Assert.DoesNotContain("Sourcefirst", text);
    }

    [Fact]
    public void InsertPageRange_ImportsInclusiveSourceRangeBeforeTargetPage() {
        byte[] target = BuildPdf("Target", "Target first", "Target second", "Target third");
        byte[] source = BuildPdf("Source", "Source first", "Source second", "Source third", "Source fourth");

        byte[] imported = PdfPageImporter.InsertPageRange(target, source, 3, 2, 3);

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(5, info.PageCount);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetfirst", "Targetsecond", "Sourcesecond", "Sourcethird", "Targetthird");
        Assert.DoesNotContain("Sourcefirst", text);
        Assert.DoesNotContain("Sourcefourth", text);
    }

    [Fact]
    public void InsertPageRange_AcceptsPdfPageRange() {
        byte[] target = BuildPdf("Target", "Target first", "Target second", "Target third");
        byte[] source = BuildPdf("Source", "Source first", "Source second", "Source third", "Source fourth");

        byte[] imported = PdfPageImporter.InsertPageRange(target, source, 3, PdfPageRange.From(2, 3));

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(5, info.PageCount);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetfirst", "Targetsecond", "Sourcesecond", "Sourcethird", "Targetthird");
        Assert.DoesNotContain("Sourcefirst", text);
        Assert.DoesNotContain("Sourcefourth", text);
    }

    [Fact]
    public void AppendPageRanges_ImportsParsedRangesInCallerOrderAndClonesOverlap() {
        byte[] target = BuildPdf("Target", "Target page");
        byte[] source = BuildPdf("Source", "Source first", "Source second", "Source third", "Source fourth");

        byte[] imported = PdfPageImporter.AppendPageRanges(target, source, PdfPageRange.ParseMany("2-3,3,1"));

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(5, info.PageCount);
        Assert.Equal("Target", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetpage", "Sourcesecond", "Sourcethird", "Sourcethird", "Sourcefirst");
        Assert.Equal(2, CountOccurrences(text, "Sourcethird"));
        Assert.DoesNotContain("Sourcefourth", text);
    }

    [Fact]
    public void PrependPageRanges_ImportsParsedRangesBeforeTarget() {
        byte[] target = BuildPdf("Target", "Target first", "Target second");
        byte[] source = BuildPdf("Source", "Source first", "Source second", "Source third");

        byte[] imported = PdfPageImporter.PrependPageRanges(target, source, PdfPageRange.ParseMany("3,1-2"));

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(5, info.PageCount);
        Assert.Equal("Target", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Sourcethird", "Sourcefirst", "Sourcesecond", "Targetfirst", "Targetsecond");
    }

    [Fact]
    public void InsertPageRanges_ImportsParsedRangesBeforeTargetPage() {
        byte[] target = BuildPdf("Target", "Target first", "Target second", "Target third");
        byte[] source = BuildPdf("Source", "Source first", "Source second", "Source third", "Source fourth");

        byte[] imported = PdfPageImporter.InsertPageRanges(target, source, 3, PdfPageRange.ParseMany("4,1-2"));

        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(6, info.PageCount);
        Assert.Equal("Target", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetfirst", "Targetsecond", "Sourcefourth", "Sourcefirst", "Sourcesecond", "Targetthird");
        Assert.DoesNotContain("Sourcethird", text);
    }

    [Fact]
    public void ImportPages_ReadsStreamsFromCurrentPositionsAndWritesOutputStreamAtCurrentPosition() {
        using var target = CreatePrefixedStream(BuildPdf("Target stream", "Target stream page"));
        using var source = CreatePrefixedStream(BuildPdf("Source stream", "Source stream first", "Source stream second"));
        using var output = CreateOutputStream(out int prefixLength);

        PdfPageImporter.AppendPages(target, source, output, 2);

        byte[] imported = GetOutputPayload(output, prefixLength);
        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(2, info.PageCount);
        Assert.Equal("Target stream", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetstreampage", "Sourcestreamsecond");
        Assert.DoesNotContain("Sourcestreamfirst", text);
    }

    [Fact]
    public void InsertPages_ReadsStreamsFromCurrentPositionsAndWritesOutputStreamAtCurrentPosition() {
        using var target = CreatePrefixedStream(BuildPdf("Target stream", "Target stream first", "Target stream second"));
        using var source = CreatePrefixedStream(BuildPdf("Source stream", "Source stream first", "Source stream second", "Source stream third"));
        using var output = CreateOutputStream(out int prefixLength);

        PdfPageImporter.InsertPageRanges(target, source, output, 2, PdfPageRange.From(2, 3));

        byte[] imported = GetOutputPayload(output, prefixLength);
        PdfDocumentInfo info = PdfInspector.Inspect(imported);
        Assert.Equal(4, info.PageCount);
        Assert.Equal("Target stream", info.Metadata.Title);

        string text = NormalizeExtractedText(PdfReadDocument.Open(imported).ExtractText());
        AssertContainsInOrder(text, "Targetstreamfirst", "Sourcestreamsecond", "Sourcestreamthird", "Targetstreamsecond");
        Assert.DoesNotContain("Sourcestreamfirst", text);
    }

    [Fact]
    public void ImportPathInputs_ReturnBytesAndCanWriteOutputForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-import-" + Guid.NewGuid().ToString("N"));
        string targetPath = Path.Combine(directory, "target.pdf");
        string sourcePath = Path.Combine(directory, "source.pdf");
        string outputPath = Path.Combine(directory, "out", "imported.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(targetPath, BuildPdf("Target path", "Target path page"));
            File.WriteAllBytes(sourcePath, BuildPdf("Source path", "Source path first", "Source path second", "Source path third"));

            byte[] prepended = PdfPageImporter.PrependPages(targetPath, sourcePath, 2);
            string prependedText = NormalizeExtractedText(PdfReadDocument.Open(prepended).ExtractText());
            AssertContainsInOrder(prependedText, "Sourcepathsecond", "Targetpathpage");
            Assert.DoesNotContain("Sourcepathfirst", prependedText);

            PdfPageImporter.AppendPages(targetPath, sourcePath, outputPath, 1, 3);

            Assert.True(File.Exists(outputPath));
            string outputText = NormalizeExtractedText(PdfReadDocument.Open(outputPath).ExtractText());
            AssertContainsInOrder(outputText, "Targetpathpage", "Sourcepathfirst", "Sourcepaththird");
            Assert.DoesNotContain("Sourcepathsecond", outputText);

            byte[] inserted = PdfPageImporter.InsertPages(targetPath, sourcePath, 1, 3);
            Assert.Equal("Target path", PdfInspector.Inspect(inserted).Metadata.Title);
            string insertedText = NormalizeExtractedText(PdfReadDocument.Open(inserted).ExtractText());
            AssertContainsInOrder(insertedText, "Sourcepaththird", "Targetpathpage");
            Assert.DoesNotContain("Sourcepathfirst", insertedText);

            string rangeOutputPath = Path.Combine(directory, "out", "inserted-range.pdf");
            PdfPageImporter.InsertPageRange(targetPath, sourcePath, rangeOutputPath, 2, 2, 3);

            Assert.True(File.Exists(rangeOutputPath));
            string rangeOutputText = NormalizeExtractedText(PdfReadDocument.Open(rangeOutputPath).ExtractText());
            AssertContainsInOrder(rangeOutputText, "Targetpathpage", "Sourcepathsecond", "Sourcepaththird");
            Assert.DoesNotContain("Sourcepathfirst", rangeOutputText);

            byte[] rangeListPrepended = PdfPageImporter.PrependPageRanges(targetPath, sourcePath, PdfPageRange.ParseMany("3,1"));
            string rangeListPrependedText = NormalizeExtractedText(PdfReadDocument.Open(rangeListPrepended).ExtractText());
            AssertContainsInOrder(rangeListPrependedText, "Sourcepaththird", "Sourcepathfirst", "Targetpathpage");
            Assert.DoesNotContain("Sourcepathsecond", rangeListPrependedText);

            string rangeListOutputPath = Path.Combine(directory, "out", "inserted-range-list.pdf");
            PdfPageImporter.InsertPageRanges(targetPath, sourcePath, rangeListOutputPath, 1, PdfPageRange.ParseMany("2-3,3"));

            Assert.True(File.Exists(rangeListOutputPath));
            string rangeListOutputText = NormalizeExtractedText(PdfReadDocument.Open(rangeListOutputPath).ExtractText());
            AssertContainsInOrder(rangeListOutputText, "Sourcepathsecond", "Sourcepaththird", "Sourcepaththird", "Targetpathpage");
            Assert.Equal(2, CountOccurrences(rangeListOutputText, "Sourcepaththird"));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ImportPathInputs_WriteToOutputStreamsForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-import-stream-" + Guid.NewGuid().ToString("N"));
        string targetPath = Path.Combine(directory, "target.pdf");
        string sourcePath = Path.Combine(directory, "source.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(targetPath, BuildPdf("Target path stream", "Target stream page"));
            File.WriteAllBytes(sourcePath, BuildPdf("Source path stream", "Source first", "Source second", "Source third", "Source fourth"));

            using var appended = CreateOutputStream(out int appendedPrefixLength);
            PdfPageImporter.AppendPages(targetPath, sourcePath, appended, 4, 1);
            string appendedText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(appended, appendedPrefixLength)).ExtractText());
            AssertContainsInOrder(appendedText, "Targetstreampage", "Sourcefourth", "Sourcefirst");
            Assert.DoesNotContain("Sourcesecond", appendedText);

            using var prepended = CreateOutputStream(out int prependedPrefixLength);
            PdfPageImporter.PrependPages(targetPath, sourcePath, prepended, 2);
            string prependedText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(prepended, prependedPrefixLength)).ExtractText());
            AssertContainsInOrder(prependedText, "Sourcesecond", "Targetstreampage");
            Assert.DoesNotContain("Sourcefirst", prependedText);

            using var inserted = CreateOutputStream(out int insertedPrefixLength);
            PdfPageImporter.InsertPages(targetPath, sourcePath, inserted, 1, 3);
            byte[] insertedBytes = GetOutputPayload(inserted, insertedPrefixLength);
            Assert.Equal("Target path stream", PdfInspector.Inspect(insertedBytes).Metadata.Title);
            string insertedText = NormalizeExtractedText(PdfReadDocument.Open(insertedBytes).ExtractText());
            AssertContainsInOrder(insertedText, "Sourcethird", "Targetstreampage");
            Assert.DoesNotContain("Sourcefirst", insertedText);

            using var integerRange = CreateOutputStream(out int integerRangePrefixLength);
            PdfPageImporter.InsertPageRange(targetPath, sourcePath, integerRange, 2, 2, 3);
            string integerRangeText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(integerRange, integerRangePrefixLength)).ExtractText());
            AssertContainsInOrder(integerRangeText, "Targetstreampage", "Sourcesecond", "Sourcethird");
            Assert.DoesNotContain("Sourcefourth", integerRangeText);

            using var pdfPageRange = CreateOutputStream(out int pdfPageRangePrefixLength);
            PdfPageImporter.InsertPageRange(targetPath, sourcePath, pdfPageRange, 2, PdfPageRange.From(1, 2));
            string pdfPageRangeText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(pdfPageRange, pdfPageRangePrefixLength)).ExtractText());
            AssertContainsInOrder(pdfPageRangeText, "Targetstreampage", "Sourcefirst", "Sourcesecond");
            Assert.DoesNotContain("Sourcethird", pdfPageRangeText);

            using var appendedRanges = CreateOutputStream(out int appendedRangesPrefixLength);
            PdfPageImporter.AppendPageRanges(targetPath, sourcePath, appendedRanges, PdfPageRange.ParseMany("2-3,3"));
            string appendedRangesText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(appendedRanges, appendedRangesPrefixLength)).ExtractText());
            AssertContainsInOrder(appendedRangesText, "Targetstreampage", "Sourcesecond", "Sourcethird", "Sourcethird");
            Assert.Equal(2, CountOccurrences(appendedRangesText, "Sourcethird"));

            using var prependedRanges = CreateOutputStream(out int prependedRangesPrefixLength);
            PdfPageImporter.PrependPageRanges(targetPath, sourcePath, prependedRanges, PdfPageRange.ParseMany("4,1"));
            string prependedRangesText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(prependedRanges, prependedRangesPrefixLength)).ExtractText());
            AssertContainsInOrder(prependedRangesText, "Sourcefourth", "Sourcefirst", "Targetstreampage");
            Assert.DoesNotContain("Sourcesecond", prependedRangesText);

            using var insertedRanges = CreateOutputStream(out int insertedRangesPrefixLength);
            PdfPageImporter.InsertPageRanges(targetPath, sourcePath, insertedRanges, 1, PdfPageRange.ParseMany("3-4"));
            string insertedRangesText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(insertedRanges, insertedRangesPrefixLength)).ExtractText());
            AssertContainsInOrder(insertedRangesText, "Sourcethird", "Sourcefourth", "Targetstreampage");
            Assert.DoesNotContain("Sourcefirst", insertedRangesText);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void ImportPages_RejectsInvalidInputs() {
        byte[] target = BuildPdf("Target", "Target page");
        byte[] source = BuildPdf("Source", "Source page");

        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPages((byte[])null!, source));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPages(target, (byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPages(target, source, (int[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.PrependPages((Stream)null!, new MemoryStream(source)));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.PrependPages(new MemoryStream(target), (Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.AppendPages(new WriteOnlyStream(), new MemoryStream(source)));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.AppendPages(new MemoryStream(target), new WriteOnlyStream()));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPages(target, source, null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.PrependPages(target, source, new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImporter.AppendPages(target, source, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages((byte[])null!, source, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages(target, (byte[])null!, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages(target, source, 1, (int[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPageRanges((byte[])null!, source, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPageRanges(target, (byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPageRanges(target, source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.AppendPageRanges(target, source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImporter.AppendPageRanges(target, source, PdfPageRange.From(1, 2)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImporter.InsertPages(target, source, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImporter.InsertPages(target, source, 3, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImporter.InsertPageRange(target, source, 1, 2, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImporter.InsertPageRange(target, source, 1, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPageRanges(target, source, 1, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPageRanges(target, source, 1, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageImporter.InsertPageRanges(target, source, 1, PdfPageRange.From(1, 2)));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages((Stream)null!, new MemoryStream(source), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages(new MemoryStream(target), (Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPages(new WriteOnlyStream(), new MemoryStream(source), 1));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPages(new MemoryStream(target), new WriteOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages(target, source, null!, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPageRange(target, source, new ReadOnlyStream(), 1, 1, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPageRanges((Stream)null!, new MemoryStream(source), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPageRanges(new MemoryStream(target), (Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.AppendPageRanges(new WriteOnlyStream(), new MemoryStream(source), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.AppendPageRanges(new MemoryStream(target), new WriteOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPageRanges(target, source, null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPageRanges(target, source, new ReadOnlyStream(), 1, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPages((string)null!, "source.pdf"));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.AppendPages(" ", "source.pdf"));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPages("target.pdf", (string)null!));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.AppendPages("target.pdf", " "));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPages("target.pdf", "source.pdf", (int[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPageRanges("target.pdf", "source.pdf", (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.AppendPages("target.pdf", "source.pdf", " "));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.AppendPages("target.pdf", "source.pdf", (Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.PrependPages("target.pdf", "source.pdf", new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages((string)null!, "source.pdf", 1));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPages(" ", "source.pdf", 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages("target.pdf", (string)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPages("target.pdf", " ", 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages("target.pdf", "source.pdf", 1, (int[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPageRanges("target.pdf", "source.pdf", 1, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPages("target.pdf", "source.pdf", " ", 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageImporter.InsertPages("target.pdf", "source.pdf", (Stream)null!, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageImporter.InsertPageRanges("target.pdf", "source.pdf", new ReadOnlyStream(), 1, PdfPageRange.From(1, 1)));
    }

    [Fact]
    public void ImportPages_RejectsDirectoryOutputTargetsBeforeReadingInputs() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-import-output-" + Guid.NewGuid().ToString("N"));
        string targetPath = Path.Combine(directory, "missing-target.pdf");
        string sourcePath = Path.Combine(directory, "missing-source.pdf");
        string outputDirectory = Path.Combine(directory, "existing-output");

        try {
            Directory.CreateDirectory(outputDirectory);

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfPageImporter.AppendPages(targetPath, sourcePath, outputDirectory, 1));

            Assert.Equal("outputPath", exception.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", exception.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    private static byte[] BuildPdf(string title, params string[] pages) {
        var doc = PdfDocument.Create()
            .Meta(title: title, author: "OfficeIMO")
            .Paragraph(p => p.Text(pages[0]));

        if (pages.Length > 1) {
            doc.Compose(compose => {
                for (int i = 1; i < pages.Length; i++) {
                    string text = pages[i];
                    compose.Page(page =>
                        page.Content(content =>
                            content.Column(column =>
                                column.Item().Paragraph(p => p.Text(text)))));
                }
            });
        }

        return doc.ToBytes();
    }

    private static byte[] BuildAnnotatedPdf(string title, string pageText) {
        return PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Meta(title: title, author: "OfficeIMO")
            .Paragraph(p => p.Text(pageText))
            .FreeTextAnnotation(
                "Import review note",
                width: 150,
                height: 44,
                borderColor: new PdfColor(0.2D, 0.4D, 0.8D),
                fillColor: new PdfColor(0.95D, 0.98D, 1D))
            .HighlightAnnotation("Import highlight", width: 120, height: 14, color: new PdfColor(1D, 0.9D, 0.1D))
            .ToBytes();
    }

    private static byte[] BuildTwoPageLabelPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 7 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
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
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 6 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Nums [0 << /S /D /St 1 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static void AssertContainsInOrder(string text, params string[] expected) {
        int previous = -1;
        foreach (string item in expected) {
            int index = text.IndexOf(item, previous + 1, StringComparison.Ordinal);
            Assert.True(index >= 0, "Expected text '" + item + "' was not found after index " + previous + " in '" + text + "'.");
            previous = index;
        }
    }

    private static string NormalizeExtractedText(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static MemoryStream CreatePrefixedStream(byte[] pdf) {
        byte[] prefix = Encoding.ASCII.GetBytes("prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        stream.Write(pdf, 0, pdf.Length);
        stream.Position = prefix.Length;
        return stream;
    }

    private static MemoryStream CreateOutputStream(out int prefixLength) {
        byte[] prefix = Encoding.ASCII.GetBytes("output-prefix");
        var stream = new MemoryStream();
        stream.Write(prefix, 0, prefix.Length);
        prefixLength = prefix.Length;
        return stream;
    }

    private static byte[] GetOutputPayload(MemoryStream output, int prefixLength) {
        byte[] bytes = output.ToArray();
        Assert.True(bytes.Length > prefixLength);
        Assert.Equal("output-prefix", Encoding.ASCII.GetString(bytes, 0, prefixLength));

        var payload = new byte[bytes.Length - prefixLength];
        Array.Copy(bytes, prefixLength, payload, 0, payload.Length);
        return payload;
    }

    private sealed class WriteOnlyStream : MemoryStream {
        public override bool CanRead => false;
    }

    private sealed class ReadOnlyStream : MemoryStream {
        public override bool CanWrite => false;
    }
}
