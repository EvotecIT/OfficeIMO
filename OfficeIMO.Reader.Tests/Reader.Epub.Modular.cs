using OfficeIMO.Epub;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

internal static class ReaderCurrentDirectoryLock {
    internal static readonly object Gate = new();
}

public sealed partial class ReaderEpubModularTests {
    [Fact]
    public async Task EpubDocument_LoadAsync_MatchesSynchronousFileSharing() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-shared-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            using var producer = new FileStream(
                epubPath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.ReadWrite | FileShare.Delete);

            EpubDocument document = await EpubDocument.LoadAsync(epubPath);

            Assert.Equal("Demo Book", document.Title);
            Assert.Equal(2, document.Chapters.Count);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_RichDispatch_MapsChaptersTablesLinksAndManifestAssets() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddEpubHandler().Build();

            OfficeDocumentReadResult result = reader.ReadDocument(epubPath);

            Assert.Equal(ReaderInputKind.Epub, result.Kind);
            Assert.Equal("Demo Book", result.Source.Title);
            Assert.Equal("OfficeIMO Team", result.Source.Author);
            Assert.Equal(2, result.Pages.Count);
            Assert.Contains(result.Blocks, block => block.Kind == "heading" && block.Text == "Two");
            Assert.Contains(result.Tables, table => table.Kind == "html-table" && table.Rows.Any(row => row.Contains("2")));
            Assert.Contains(result.Links, link => link.Uri == "https://example.test/chapter-two" && link.Text == "details");
            string markdown = Assert.IsType<string>(result.Markdown);
            Assert.Contains("## Second", markdown, StringComparison.Ordinal);
            Assert.Contains("# Two", markdown, StringComparison.Ordinal);
            Assert.Contains("- EPUB list item", markdown, StringComparison.Ordinal);
            Assert.Contains("[details](https://example.test/chapter-two)", markdown, StringComparison.Ordinal);
            Assert.All(result.Chunks, chunk => Assert.Equal(ReaderInputKind.Epub, chunk.Kind));
            Assert.Equal(
                result.Source.Path + "::OEBPS/chapter2.xhtml#details",
                Assert.Single(result.Links, link => link.Text == "next chapter").Uri);
            OfficeDocumentAsset asset = Assert.Single(result.Assets, item => item.MediaType == "image/png");
            Assert.NotNull(asset.PayloadBytes);
            Assert.True(asset.PayloadHashMatches(out _));
            OfficeDocumentPage coverPage = Assert.Single(result.Pages, page => page.Location.Path?.EndsWith("::OEBPS/chapter2.xhtml", StringComparison.Ordinal) == true);
            Assert.Contains(coverPage.Assets, pageAsset => ReferenceEquals(pageAsset, asset));
            OfficeDocumentAsset[] inlineImages = result.Assets.Where(item => item.MediaType == "image/gif").ToArray();
            Assert.Equal(2, inlineImages.Length);
            Assert.Equal(2, inlineImages.Select(item => item.FileName).Distinct(StringComparer.Ordinal).Count());
            Assert.All(result.Pages, page => Assert.Contains(page.Assets, pageAsset => pageAsset.MediaType == "image/gif"));
            ReaderVisual visual = Assert.Single(result.Visuals, item =>
                item.Kind == "image" &&
                item.SourceName == result.Source.Path + "::OEBPS/images/cover.png");
            Assert.StartsWith("epub-chapter-0001-html-image-", visual.Location!.BlockAnchor!, StringComparison.Ordinal);
            using (FileStream stream = File.OpenRead(epubPath)) {
                OfficeDocumentReadResult jsonResult = OfficeDocumentReadResultJson.Deserialize(
                    EpubReaderAdapter.ReadDocumentJson(stream, "book.epub"));
                Assert.Equal(ReaderInputKind.Epub, jsonResult.Kind);
            }
            Assert.Contains("officeimo.reader.epub.rich-v5", result.CapabilitiesUsed);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_BuilderPreservesRegisteredRawHtmlBudget() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            var registeredOptions = new EpubReadOptions { MaxTotalRawHtmlBytes = 1 };
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddEpubHandler(registeredOptions)
                .Build();
            registeredOptions.MaxTotalRawHtmlBytes = long.MaxValue;

            OfficeDocumentReadResult result = reader.ReadDocument(epubPath);

            Assert.Contains(result.Chunks, chunk =>
                chunk.Warnings?.Any(warning => warning.Contains("MaxTotalRawHtmlBytes", StringComparison.Ordinal)) == true);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_DirectRead_PreservesStructuredMarkdownAndChapterProvenanceByDefault() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            var epubOptions = new EpubReadOptions {
                IncludeRawHtml = false,
                PreferSpineOrder = true
            };

            ReaderChunk[] chunks = EpubReaderAdapter.Read(
                epubPath,
                readerOptions: new ReaderOptions { MaxChars = 4_000 },
                epubOptions: epubOptions).ToArray();

            ReaderChunk secondChapter = Assert.Single(
                chunks,
                chunk => chunk.Location.Path?.Contains("::OEBPS/chapter2.xhtml", StringComparison.OrdinalIgnoreCase) == true);
            ReaderChunk firstChapter = Assert.Single(
                chunks,
                chunk => chunk.Location.Path?.Contains("::OEBPS/chapter1.xhtml", StringComparison.OrdinalIgnoreCase) == true);
            Assert.Contains("## Second", secondChapter.Markdown, StringComparison.Ordinal);
            Assert.Contains("# Two", secondChapter.Markdown, StringComparison.Ordinal);
            Assert.Contains("- EPUB list item", secondChapter.Markdown, StringComparison.Ordinal);
            Assert.Contains("[details](https://example.test/chapter-two)", secondChapter.Markdown, StringComparison.Ordinal);
            Assert.Equal(0, secondChapter.Location.SourceBlockIndex);
            Assert.Equal(1, firstChapter.Location.SourceBlockIndex);
            Assert.Equal("Second", secondChapter.Location.HeadingPath);
            Assert.Equal(
                new[] { "Second", "Two" },
                ReaderHeadingPath.Split(secondChapter.Location.HierarchyHeadingPath));
            Assert.False(epubOptions.IncludeRawHtml);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_RichDispatch_ProjectsPackageLayoutEncryptionAndDiagnostics() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath, includePackageDiagnostics: true);

            OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(epubPath);

            Assert.Equal("3.0", Assert.Single(result.Metadata, item => item.Id == "epub-package-version").Value);
            Assert.Equal("PrePaginated", Assert.Single(result.Metadata, item => item.Id == "epub-rendition-layout").Value);
            Assert.Equal("True", Assert.Single(result.Metadata, item => item.Id == "epub-fixed-layout").Value);
            Assert.Equal("1", Assert.Single(result.Metadata, item => item.Id == "epub-encryption-count").Value);
            Assert.Equal("True", Assert.Single(result.Metadata, item => item.Id == "epub-requires-decryption").Value);
            OfficeDocumentDiagnostic encryption = Assert.Single(result.Diagnostics, item => item.Code == "epub.encryption.unsupported");
            Assert.Equal(OfficeDocumentDiagnosticCategory.Security, encryption.Category);
            Assert.EndsWith("::OEBPS/protected.bin", encryption.Location!.Path, StringComparison.Ordinal);
            Assert.DoesNotContain(result.Diagnostics, item =>
                item.Code == "reader-warning" &&
                string.Equals(item.Message, encryption.Message, StringComparison.Ordinal));
            OfficeDocumentDiagnostic layout = Assert.Single(result.Diagnostics, item => item.Code == "epub.layout.fixed");
            Assert.Equal(OfficeDocumentDiagnosticCategory.Content, layout.Category);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_RichDispatch_MapsEncryptedChapterDiagnosticsToSecurity() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-encrypted-chapter-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithEncryptedChapter(epubPath);

            OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(epubPath);

            OfficeDocumentDiagnostic diagnostic = Assert.Single(
                result.Diagnostics,
                item => item.Code == "epub.chapter.encrypted");
            Assert.Equal(OfficeDocumentDiagnosticCategory.Security, diagnostic.Category);
            Assert.EndsWith("::OEBPS/locked.xhtml", diagnostic.Location!.Path, StringComparison.Ordinal);
            OfficeDocumentPage page = Assert.Single(result.Pages);
            Assert.EndsWith("::OEBPS/open.xhtml", page.Location.Path, StringComparison.Ordinal);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_RichDispatch_HonorsDisabledResourcePayloadsAndKeepsPageAssets() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);

            OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(
                epubPath,
                epubOptions: new EpubReadOptions { IncludeResourceData = false });

            OfficeDocumentAsset asset = Assert.Single(result.Assets, item => item.MediaType == "image/png");
            Assert.Null(asset.PayloadBytes);
            Assert.Null(asset.PayloadHash);
            OfficeDocumentPage coverPage = Assert.Single(result.Pages, page => page.Location.Path?.EndsWith("::OEBPS/chapter2.xhtml", StringComparison.Ordinal) == true);
            Assert.Contains(coverPage.Assets, pageAsset => ReferenceEquals(pageAsset, asset));
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_RichDispatch_PreservesImageOnlySpinePagesAndPageOcrCandidates() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-image-only-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildImageOnlyEpub(epubPath);

            EpubDocument document = EpubDocument.Load(epubPath, new EpubReadOptions { IncludeRawHtml = true });
            EpubChapter chapter = Assert.Single(document.Chapters);
            Assert.Equal(string.Empty, chapter.Text);
            Assert.NotNull(chapter.Html);

            OfficeDocumentReadResult result = EpubReaderAdapter.ReadDocument(epubPath);

            OfficeDocumentPage page = Assert.Single(result.Pages);
            OfficeDocumentAsset asset = Assert.Single(page.Assets);
            Assert.Same(asset, Assert.Single(result.Assets));
            OfficeDocumentOcrCandidate candidate = Assert.Single(result.OcrCandidates);
            Assert.Equal(asset.Id, candidate.AssetId);
            Assert.Same(candidate, Assert.Single(page.OcrCandidates));
            Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "ocr-needed");
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void EpubReader_UsesOpfSpineOrderAndMetadata() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);

            var document = EpubDocument.Load(epubPath, new EpubReadOptions {
                PreferSpineOrder = true,
                IncludeNonLinearSpineItems = true
            });

            Assert.NotNull(document);
            Assert.Equal("Demo Book", document.Title);
            Assert.Equal("author-1", document.Identifier);
            Assert.Equal("en", document.Language);
            Assert.Equal("OfficeIMO Team", document.Creator);
            Assert.Equal("OEBPS/content.opf", document.OpfPath);
            Assert.Equal(2, document.Chapters.Count);

            var first = document.Chapters[0];
            var second = document.Chapters[1];

            Assert.Equal("OEBPS/chapter2.xhtml", first.Path);
            Assert.Equal("Second", first.Title);
            Assert.Equal(1, first.SpineIndex);
            Assert.True(first.IsLinear);
            Assert.Contains("Second chapter text.", first.Text, StringComparison.Ordinal);

            Assert.Equal("OEBPS/chapter1.xhtml", second.Path);
            Assert.Equal("First", second.Title);
            Assert.Equal(2, second.SpineIndex);
            Assert.True(second.IsLinear);
            Assert.Contains("First chapter text.", second.Text, StringComparison.Ordinal);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void EpubReader_ManifestResourcePayloads_AreOptInAndBounded() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);

            EpubDocument metadataOnly = EpubDocument.Load(epubPath);
            EpubResource imageMetadata = Assert.Single(metadataOnly.Resources, resource => resource.MediaType == "image/png");
            Assert.Null(imageMetadata.Data);

            EpubDocument bounded = EpubDocument.Load(epubPath, new EpubReadOptions {
                IncludeResourceData = true,
                MaxResourceBytes = 4,
                MaxTotalResourceBytes = 32
            });
            EpubResource boundedImage = Assert.Single(bounded.Resources, resource => resource.MediaType == "image/png");
            Assert.Null(boundedImage.Data);
            Assert.Contains(bounded.Warnings, warning => warning.Contains("MaxResourceBytes", StringComparison.Ordinal));

            EpubDocument withPayload = EpubDocument.Load(epubPath, new EpubReadOptions { IncludeResourceData = true });
            EpubResource image = Assert.Single(withPayload.Resources, resource => resource.MediaType == "image/png");
            byte[] firstRead = Assert.IsType<byte[]>(image.Data);
            byte original = firstRead[0];
            firstRead[0] ^= byte.MaxValue;
            Assert.Equal(original, image.Data![0]);

            Assert.False(typeof(EpubDocument).GetProperty(nameof(EpubDocument.Title))!.SetMethod!.IsPublic);
            Assert.False(typeof(EpubChapter).GetProperty(nameof(EpubChapter.Text))!.SetMethod!.IsPublic);
            Assert.False(typeof(EpubResource).GetProperty(nameof(EpubResource.Data))!.SetMethod!.IsPublic);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void EpubReader_RawHtmlRetention_IsAggregateBounded() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);

            EpubDocument document = EpubDocument.Load(epubPath, new EpubReadOptions {
                IncludeRawHtml = true,
                MaxTotalRawHtmlBytes = 1
            });

            Assert.Equal(2, document.Chapters.Count);
            Assert.All(document.Chapters, chapter => Assert.Null(chapter.Html));
            Assert.All(document.Chapters, chapter => Assert.False(string.IsNullOrWhiteSpace(chapter.Text)));
            Assert.Contains(document.Diagnostics, diagnostic =>
                diagnostic.Code == "epub.chapter.raw-html-total-limit" &&
                diagnostic.Severity == EpubDiagnosticSeverity.Warning);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_EmitsWarningsAndVirtualChapterPaths() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithMalformedChapter(epubPath);

            var chunks = EpubReaderAdapter.Read(
                epubPath,
                readerOptions: new ReaderOptions { MaxChars = 64 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true, FallbackToHtmlScan = true }).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Epub &&
                (c.Warnings?.Any(w => w.Contains("not valid XML", StringComparison.OrdinalIgnoreCase)) ?? false));

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Epub &&
                (c.Location.Path?.Contains("::OEBPS/good.xhtml", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("Good chapter body text.", StringComparison.Ordinal) ?? false));
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_ReadsFromNonSeekableStream() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            var bytes = File.ReadAllBytes(epubPath);

            using var stream = new NonSeekableReadStream(bytes);
            var chunks = EpubReaderAdapter.Read(
                stream,
                sourceName: "nonseekable.epub",
                readerOptions: new ReaderOptions { MaxChars = 4_000 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true }).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Epub &&
                (c.Location.Path?.Contains("nonseekable.epub::OEBPS/chapter2.xhtml", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("Second chapter text.", StringComparison.Ordinal) ?? false));
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_ReadsFromNonSeekableStream_EnforcesMaxInputBytes() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            var bytes = File.ReadAllBytes(epubPath);

            using var stream = new NonSeekableReadStream(bytes);
            var ex = Assert.Throws<IOException>(() => EpubReaderAdapter.Read(
                stream,
                sourceName: "nonseekable.epub",
                readerOptions: new ReaderOptions { MaxInputBytes = 16, MaxChars = 4_000 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true }).ToList());

            Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_UsesMonotonicBlockIndexesAcrossWarningsAndContent() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithMalformedChapter(epubPath);

            var chunks = EpubReaderAdapter.Read(
                epubPath,
                readerOptions: new ReaderOptions { MaxChars = 64 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true, FallbackToHtmlScan = true }).ToList();

            Assert.NotEmpty(chunks);

            var blockIndexes = chunks
                .Select(c => c.Location.BlockIndex)
                .ToList();

            Assert.DoesNotContain(blockIndexes, static index => !index.HasValue);
            Assert.Equal(Enumerable.Range(0, blockIndexes.Count), blockIndexes.Select(i => i!.Value));

            var contentChunk = Assert.Single(chunks, c =>
                c.Kind == ReaderInputKind.Epub &&
                (c.Location.Path?.Contains("::OEBPS/good.xhtml", StringComparison.OrdinalIgnoreCase) ?? false));

            Assert.Equal(0, contentChunk.Location.SourceBlockIndex);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_WarningChunks_UseVirtualEntryPaths() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithMalformedChapter(epubPath);

            var chunks = EpubReaderAdapter.Read(
                epubPath,
                readerOptions: new ReaderOptions { ComputeHashes = true, MaxChars = 64 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true, FallbackToHtmlScan = true }).ToList();

            var warningChunk = Assert.Single(chunks, c => c.Warnings?.Count > 0);
            var contentChunk = Assert.Single(chunks, c =>
                c.Kind == ReaderInputKind.Epub &&
                (c.Location.Path?.Contains("::OEBPS/good.xhtml", StringComparison.OrdinalIgnoreCase) ?? false));

            Assert.Contains("::OEBPS/bad.xhtml", warningChunk.Location.Path, StringComparison.OrdinalIgnoreCase);
            Assert.NotEqual(warningChunk.SourceId, contentChunk.SourceId);
            Assert.NotEqual(warningChunk.ChunkHash, contentChunk.ChunkHash);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_StreamWarningChunks_UseTrimmedVirtualEntryPaths() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithMalformedChapter(epubPath);
            var bytes = File.ReadAllBytes(epubPath);

            using var stream = new NonSeekableReadStream(bytes);
            var warningChunk = Assert.Single(
                EpubReaderAdapter.Read(
                    stream,
                    sourceName: " malformed.epub ",
                    readerOptions: new ReaderOptions { MaxChars = 64 },
                    epubOptions: new EpubReadOptions { PreferSpineOrder = true, FallbackToHtmlScan = true }),
                c => c.Warnings?.Count > 0);

            Assert.Contains("malformed.epub::OEBPS/bad.xhtml", warningChunk.Location.Path, StringComparison.OrdinalIgnoreCase);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_DirectRead_EmitsSourceAndChunkMetadata() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);

            var chunks = EpubReaderAdapter.Read(
                epubPath,
                readerOptions: new ReaderOptions { ComputeHashes = true, MaxChars = 4_000 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true }).ToList();
            string archiveHash = ComputeSha256Hex(File.ReadAllBytes(epubPath));

            Assert.NotEmpty(chunks);
            Assert.All(chunks, chunk => {
                Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(chunk.SourceHash));
                Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
                Assert.True(chunk.TokenEstimate.HasValue && chunk.TokenEstimate.Value >= 1);
                Assert.True(chunk.SourceLengthBytes.HasValue && chunk.SourceLengthBytes.Value > 0);
                Assert.True(chunk.SourceLastWriteUtc.HasValue);
                Assert.Equal(archiveHash, chunk.SourceHash);
            });

            var first = Assert.Single(chunks, c => c.Location.Path?.Contains("::OEBPS/chapter2.xhtml", StringComparison.OrdinalIgnoreCase) ?? false);
            var second = Assert.Single(chunks, c => c.Location.Path?.Contains("::OEBPS/chapter1.xhtml", StringComparison.OrdinalIgnoreCase) ?? false);

            Assert.NotEqual(first.SourceId, second.SourceId);
            Assert.NotEqual(first.ChunkHash, second.ChunkHash);
            Assert.Equal(first.SourceHash, second.SourceHash);
            Assert.Equal(first.SourceLengthBytes, second.SourceLengthBytes);
            Assert.Equal(first.SourceLastWriteUtc, second.SourceLastWriteUtc);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_PreservesLiteralChapterTitleAndAllowsVirtualChapterSources() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath, "Q1 &gt; Q2\\Back");
            var chunks = EpubReaderAdapter.Read(
                epubPath,
                readerOptions: new ReaderOptions { MaxChars = 4_000 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true }).ToList();

            ReaderChunk chapter = Assert.Single(
                chunks,
                chunk => chunk.Location.Path?.Contains("::OEBPS/chapter2.xhtml", StringComparison.OrdinalIgnoreCase) ?? false);
            Assert.Equal("Q1 > Q2\\Back", chapter.Location.HeadingPath);
            Assert.NotEqual(chunks[0].SourceId, chunks[1].SourceId);

            var document = new OfficeDocumentReadResult {
                Kind = ReaderInputKind.Epub,
                Source = new OfficeDocumentSource { SourceId = "book", Path = epubPath },
                Chunks = chunks
            };
            ReaderChunkHierarchyResult hierarchy = ReaderHierarchicalChunker.Chunk(document);
            Assert.Equal(chunks.Count, hierarchy.Chunks.Count);
            Assert.Contains(
                hierarchy.Nodes,
                node => node.Kind == ReaderChunkHierarchyNodeKind.Heading &&
                        string.Equals(node.Title, "Q1 > Q2\\Back", StringComparison.Ordinal));
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_DirectStreamRead_EmitsLogicalSourceMetadata() {
        var epubPath = Path.Combine(Path.GetTempPath(), "officeimo-epub-" + Guid.NewGuid().ToString("N") + ".epub");
        try {
            BuildEpubWithSpine(epubPath);
            var bytes = File.ReadAllBytes(epubPath);

            using var stream = new NonSeekableReadStream(bytes);
            var chunks = EpubReaderAdapter.Read(
                stream,
                sourceName: " metadata.epub ",
                readerOptions: new ReaderOptions { ComputeHashes = true, MaxChars = 4_000 },
                epubOptions: new EpubReadOptions { PreferSpineOrder = true }).ToList();
            string archiveHash = ComputeSha256Hex(bytes);

            Assert.NotEmpty(chunks);
            Assert.All(chunks, chunk => {
                Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(chunk.SourceHash));
                Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
                Assert.True(chunk.TokenEstimate.HasValue && chunk.TokenEstimate.Value >= 1);
                Assert.Equal(bytes.Length, chunk.SourceLengthBytes);
                Assert.Null(chunk.SourceLastWriteUtc);
                Assert.Equal(archiveHash, chunk.SourceHash);
            });

            Assert.All(chunks, chunk => Assert.StartsWith("metadata.epub::", chunk.Location.Path, StringComparison.OrdinalIgnoreCase));

            var first = Assert.Single(chunks, c => c.Location.Path?.Contains("::OEBPS/chapter2.xhtml", StringComparison.OrdinalIgnoreCase) ?? false);
            var second = Assert.Single(chunks, c => c.Location.Path?.Contains("::OEBPS/chapter1.xhtml", StringComparison.OrdinalIgnoreCase) ?? false);

            Assert.NotEqual(first.SourceId, second.SourceId);
            Assert.NotEqual(first.ChunkHash, second.ChunkHash);
            Assert.Equal(first.SourceHash, second.SourceHash);
            Assert.Equal(first.SourceLengthBytes, second.SourceLengthBytes);
        } finally {
            if (File.Exists(epubPath)) File.Delete(epubPath);
        }
    }

    [Fact]
    public void DocumentReaderEpub_FileReads_CanonicalizeEquivalentArchivePathsForIdentity() {
        var tempDirectory = Path.Combine(Path.GetTempPath(), "officeimo-reader-epub-canonical-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDirectory);
        var epubPath = Path.Combine(tempDirectory, "canonical.epub");

        lock (ReaderCurrentDirectoryLock.Gate) {
            var originalCurrentDirectory = Environment.CurrentDirectory;
            try {
                BuildEpubWithSpine(epubPath);

                Environment.CurrentDirectory = tempDirectory;
                var relativePath = Path.GetFileName(epubPath);
                var fullPath = Path.GetFullPath(relativePath).Replace('\\', '/');

                var relativeChunk = EpubReaderAdapter.Read(
                    relativePath,
                    readerOptions: new ReaderOptions { ComputeHashes = true, MaxChars = 4_000 },
                    epubOptions: new EpubReadOptions { PreferSpineOrder = true })
                    .Single(c => c.Location.Path?.Contains("::OEBPS/chapter2.xhtml", StringComparison.OrdinalIgnoreCase) ?? false);

                var fullChunk = EpubReaderAdapter.Read(
                    epubPath,
                    readerOptions: new ReaderOptions { ComputeHashes = true, MaxChars = 4_000 },
                    epubOptions: new EpubReadOptions { PreferSpineOrder = true })
                    .Single(c => c.Location.Path?.Contains("::OEBPS/chapter2.xhtml", StringComparison.OrdinalIgnoreCase) ?? false);

                Assert.Equal(fullPath + "::OEBPS/chapter2.xhtml", relativeChunk.Location.Path);
                Assert.Equal(relativeChunk.Location.Path, fullChunk.Location.Path);
                Assert.Equal(relativeChunk.SourceId, fullChunk.SourceId);
                Assert.Equal(relativeChunk.ChunkHash, fullChunk.ChunkHash);
                Assert.Equal(relativeChunk.SourceHash, fullChunk.SourceHash);
            } finally {
                Environment.CurrentDirectory = originalCurrentDirectory;
                if (File.Exists(epubPath)) File.Delete(epubPath);
                if (Directory.Exists(tempDirectory)) Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }

}
