using OfficeIMO.Reader.IWork;
using OfficeIMO.Reader.All;
using System.IO.Compression;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderIWorkTests {
    [Theory]
    [InlineData("sample.pages", "application/vnd.apple.pages")]
    [InlineData("sample.numbers", "application/vnd.apple.numbers")]
    [InlineData("sample.key", "application/vnd.apple.keynote")]
    public void IWorkExtensionDetectionReportsItsRegisteredMediaType(
        string sourceName, string expectedMediaType) {
        ReaderDetectionResult detection = new OfficeDocumentReaderBuilder()
            .AddIWorkHandler().Build().Detect(Array.Empty<byte>(), sourceName,
                new ReaderDetectionOptions { Mode = ReaderDetectionMode.ExtensionOnly });

        Assert.Equal(ReaderInputKind.IWork, detection.Kind);
        Assert.Equal(expectedMediaType, detection.MediaType);
    }

    [Theory]
    [InlineData("nim-iwork/simple.pages", "hello pages")]
    [InlineData("nim-iwork/simple.numbers", "a")]
    [InlineData("nim-iwork/simple.key", "hello keynote")]
    public void PublicReaderProjectsIndependentIWorkCorpusForPathsAndStreams(
        string relativePath, string marker) {
        string path = Fixture(relativePath);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();

        OfficeDocumentReadResult fromPath = reader.ReadDocument(path);
        using FileStream input = File.OpenRead(path);
        OfficeDocumentReadResult fromStream = reader.ReadDocument(input, Path.GetFileName(path));

        Assert.Equal(ReaderInputKind.IWork, fromPath.Kind);
        Assert.Contains("officeimo.reader.iwork", fromPath.CapabilitiesUsed);
        Assert.Equal(fromPath.Chunks.Select(chunk => chunk.Text),
            fromStream.Chunks.Select(chunk => chunk.Text));
        Assert.Contains(marker, string.Join("\n", fromPath.Chunks.Select(chunk => chunk.Text)),
            StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(fromPath.Pages);
        Assert.Equal(OfficeDocumentPageProvenance.LogicalContainer,
            fromPath.GetPageProvenance());
    }

    [Fact]
    public void NumbersTablesRetainCachedValuesAndReportBoundedProjection() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddIWorkHandler(new ReaderIWorkOptions { MaximumTableColumns = 2 })
            .Build();

        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("numbers-parser/test-10-formulas.numbers"),
            new ReaderOptions { MaxTableRows = 2 });

        Assert.Equal(ReaderInputKind.IWork, document.Kind);
        Assert.True(document.Pages.Count >= 2);
        Assert.All(document.Tables, table => Assert.True(table.Rows.Count <= 2));
        Assert.Contains(document.Tables, table => table.Truncated);
        Assert.Contains(document.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_TABLE_TRUNCATED");
        Assert.Contains(document.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_READER_FORMULA_CACHE");
    }

    [Fact]
    public void PagesTablesAndImagesRemainAvailableWhenBodyNeedsVisualFallback() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();

        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("picodocs/sample-v14.4.pages"));

        Assert.Equal(3, document.Tables.Count);
        Assert.Equal("Feature", document.Tables[0].Rows[0][0]);
        Assert.Contains(document.Tables[0].Rows, row => row.Contains("Preserve reading order"));
        OfficeDocumentAsset image = Assert.Single(document.Assets);
        Assert.Equal("image/png", image.MediaType);
        Assert.Null(image.PayloadBytes);
        Assert.Contains(document.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
        OfficeDocumentBlock tableBlock = Assert.Single(document.Blocks,
            block => block.Kind == "table" && block.Text.Contains("Feature", StringComparison.Ordinal));
        Assert.Equal(DocumentReaderEngine.BuildRichTableText(document.Tables[0]), tableBlock.Text);
        Assert.NotEqual(document.Tables[0].ToMarkdownTable(), tableBlock.Text);
    }

    [Fact]
    public void KeynoteProjectsSlidesNotesTablesAndImagePayloads() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddIWorkHandler(new ReaderIWorkOptions { IncludeImagePayloads = true })
            .Build();

        OfficeDocumentReadResult notes = reader.ReadDocument(Fixture("nim-iwork/simple.key"));
        OfficeDocumentReadResult table = reader.ReadDocument(
            Fixture("keynotekit/tabledeck-v15.2.1.key"));
        OfficeDocumentReadResult image = reader.ReadDocument(
            Fixture("keynotekit/imagedeck-v15.2.1.key"));

        Assert.Equal(2, notes.Pages.Count);
        Assert.Contains("note text here", notes.Markdown, StringComparison.Ordinal);
        Assert.Contains(notes.Blocks, block => block.Location.SourceBlockKind == "presenter-notes"
            && block.Text.Contains("note text here", StringComparison.Ordinal));
        Assert.Equal("Product", Assert.Single(table.Tables).Columns[0]);
        Assert.NotEmpty(Assert.Single(image.Assets).PayloadBytes!);
    }

    [Fact]
    public void AllPresetRoutesIWorkFilesThroughTheAdapter() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers()
            .Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Fixture("nim-iwork/simple.pages"));

        Assert.Equal(ReaderInputKind.IWork, result.Kind);
        Assert.Contains("officeimo.reader.iwork", result.CapabilitiesUsed);
    }

    [Theory]
    [InlineData("nim-iwork/simple.pages")]
    [InlineData("nim-iwork/simple.numbers")]
    [InlineData("nim-iwork/simple.key")]
    public async Task PreferContentRetainsValidatedIWorkRoutes(string relativePath) {
        string path = Fixture(relativePath);
        var options = new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent };
        foreach (OfficeDocumentReader reader in new[] {
                     new OfficeDocumentReaderBuilder().AddIWorkHandler().Build(),
                     new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build()
                 }) {
            OfficeDocumentReadResult fromPath = reader.ReadDocument(path, options);
            Assert.Equal(ReaderInputKind.IWork, fromPath.Kind);
            Assert.DoesNotContain(fromPath.Diagnostics,
                diagnostic => diagnostic.Code == "input-kind-mismatch");
            using FileStream input = File.OpenRead(path);
            OfficeDocumentReadResult fromStream = reader.ReadDocument(input, Path.GetFileName(path), options);
            Assert.Equal(ReaderInputKind.IWork, fromStream.Kind);
            Assert.DoesNotContain(fromStream.Diagnostics,
                diagnostic => diagnostic.Code == "input-kind-mismatch");
            OfficeDocumentReadResult asyncPath = await reader.ReadDocumentAsync(path, options);
            Assert.Equal(ReaderInputKind.IWork, asyncPath.Kind);
            Assert.DoesNotContain(asyncPath.Diagnostics,
                diagnostic => diagnostic.Code == "input-kind-mismatch");
            using FileStream asyncInput = File.OpenRead(path);
            OfficeDocumentReadResult asyncStream = await reader.ReadDocumentAsync(
                asyncInput, Path.GetFileName(path), options);
            Assert.Equal(ReaderInputKind.IWork, asyncStream.Kind);
            Assert.DoesNotContain(asyncStream.Diagnostics,
                diagnostic => diagnostic.Code == "input-kind-mismatch");
        }
    }

    [Theory]
    [InlineData("nim-iwork/simple.pages", "application/vnd.apple.pages")]
    [InlineData("nim-iwork/simple.numbers", "application/vnd.apple.numbers")]
    [InlineData("nim-iwork/simple.key", "application/vnd.apple.keynote")]
    public async Task PublicContentDetectionRecognizesIWorkPackages(
        string relativePath, string mediaType) {
        string path = Fixture(relativePath);
        string sourceName = Path.GetFileName(path);
        byte[] bytes = File.ReadAllBytes(path);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();

        foreach (ReaderDetectionResult result in new[] {
                     reader.Detect(path),
                     reader.Detect(bytes, sourceName),
                     await reader.DetectAsync(bytes, sourceName)
                 }) {
            Assert.Equal(ReaderInputKind.IWork, result.Kind);
            Assert.Equal(ReaderInputKind.IWork, result.ContentKind);
            Assert.Equal(mediaType, result.MediaType);
            Assert.False(result.IsMismatch);
        }

        byte[] prefix = { 0x19, 0x27, 0x38, 0x44, 0x55 };
        using var stream = new MemoryStream(prefix.Concat(bytes).ToArray(), writable: false);
        stream.Position = prefix.Length;
        ReaderDetectionResult offsetDetection = reader.Detect(stream, sourceName);
        Assert.Equal(ReaderInputKind.IWork, offsetDetection.Kind);
        Assert.Equal(prefix.Length, stream.Position);
    }

    [Fact]
    public void ContentDetectedIWorkCanReadWithoutAnIWorkExtension() {
        byte[] bytes = File.ReadAllBytes(Fixture("nim-iwork/simple.pages"));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers().Build();
        var options = new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent };

        ReaderDetectionResult detection = reader.Detect(bytes, "renamed.zip");
        OfficeDocumentReadResult document = reader.ReadDocument(bytes, "renamed.zip", options);

        Assert.Equal(ReaderInputKind.IWork, detection.Kind);
        Assert.Equal(ReaderInputKind.IWork, document.Kind);
        Assert.Contains(document.Chunks, chunk => chunk.Text.Contains("hello pages", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void BareIndexZipDoesNotOverrideGenericZipOrOpenXmlEvidence() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers().Build();

        byte[] generic = CreateZip("Index.zip", "notes.txt");
        ReaderDetectionResult genericDetection = reader.Detect(generic, "notes.zip");
        Assert.Equal(ReaderInputKind.Zip, genericDetection.Kind);

        byte[] nestedGeneric = CreateZip("notes.txt");
        using var outerStream = new MemoryStream();
        using (var outer = new ZipArchive(outerStream, ZipArchiveMode.Create, leaveOpen: true)) {
            using Stream index = outer.CreateEntry("Index.zip").Open();
            index.Write(nestedGeneric, 0, nestedGeneric.Length);
        }
        Assert.Equal(ReaderInputKind.Zip, reader.Detect(outerStream.ToArray(), "notes.zip").Kind);

        byte[] renamedGeneric = outerStream.ToArray();
        Assert.Equal(ReaderInputKind.Zip, reader.Detect(renamedGeneric, "notes.pages").Kind);
        Assert.NotEqual(ReaderInputKind.IWork,
            reader.ReadDocument(renamedGeneric, "notes.pages",
                new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent }).Kind);

        byte[] openXml = CreateZip("Index/Document.iwa", "word/document.xml");
        ReaderDetectionResult wordDetection = reader.Detect(openXml, "document.docx");
        Assert.Equal(ReaderInputKind.Word, wordDetection.Kind);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NestedIWorkIndexIsDetectedWithoutAnIWorkExtension(bool preserveIndexPrefix) {
        byte[] nestedPackage = CreateNestedIndexPackage(
            File.ReadAllBytes(Fixture("nim-iwork/simple.pages")), preserveIndexPrefix);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers().Build();

        ReaderDetectionResult sync = reader.Detect(nestedPackage, "renamed.zip");
        ReaderDetectionResult asyncResult = await reader.DetectAsync(nestedPackage, "renamed.zip");
        OfficeDocumentReadResult document = reader.ReadDocument(nestedPackage, "renamed.zip",
            new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });
        OfficeDocumentReadResult namedDocument = reader.ReadDocument(nestedPackage,
            "renamed.pages", new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });
        OfficeDocumentReadResult asyncDocument = await reader.ReadDocumentAsync(nestedPackage,
            "renamed.zip", new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });

        Assert.Equal(ReaderInputKind.IWork, sync.Kind);
        Assert.Equal(ReaderInputKind.IWork, asyncResult.Kind);
        Assert.Equal(ReaderInputKind.IWork, document.Kind);
        Assert.Equal(ReaderInputKind.IWork, namedDocument.Kind);
        Assert.Equal(ReaderInputKind.IWork, asyncDocument.Kind);
        Assert.Equal(document.Chunks.Select(chunk => chunk.Text), asyncDocument.Chunks.Select(chunk => chunk.Text));
        Assert.Contains(document.Chunks, chunk =>
            chunk.Text.Contains("hello pages", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task NestedIWorkIndexWithZip64EntryMetadataIsDetected() {
        byte[] nestedPackage = CreateNestedIndexPackage(
            File.ReadAllBytes(Fixture("nim-iwork/simple.pages")));
        byte[] package = WrapIndexEntryWithZip64Metadata(nestedPackage);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers().Build();

        Assert.Equal(ReaderInputKind.IWork, reader.Detect(package, "renamed.zip").Kind);
        Assert.Equal(ReaderInputKind.IWork,
            (await reader.DetectAsync(package, "renamed.zip")).Kind);
        Assert.Contains(reader.ReadDocument(package, "renamed.zip",
                new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent }).Chunks,
            chunk => chunk.Text.Contains("hello pages", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task PublicDetectionRecognizesZip64IWorkPackage() {
        byte[] package = WrapWithZip64EndRecord(
            File.ReadAllBytes(Fixture("nim-iwork/simple.pages")));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();

        ReaderDetectionResult sync = reader.Detect(package, "large.pages");
        ReaderDetectionResult asyncResult = await reader.DetectAsync(package, "large.pages");
        OfficeDocumentReadResult document = reader.ReadDocument(package, "large.pages",
            new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });

        Assert.Equal(ReaderInputKind.IWork, sync.Kind);
        Assert.Equal(ReaderInputKind.IWork, asyncResult.Kind);
        Assert.Equal(ReaderInputKind.IWork, document.Kind);
    }

    [Fact]
    public void PreferContentUsesTheKnownIWorkPackageLimitBeforeDetection() {
        const long packageLimit = 512L * 1024L * 1024L;
        var registry = new ReaderHandlerRegistry();
        registry.Register(new ReaderHandlerRegistration {
            Id = "officeimo.tests.iwork-limit",
            Kind = ReaderInputKind.IWork,
            Extensions = new[] { ".pages" },
            DefaultMaxInputBytes = packageLimit,
            MaxInputBytesCeiling = packageLimit,
            ReadPath = (_, _, _) => Array.Empty<ReaderChunk>(),
            ReadStream = (_, _, _, _) => Array.Empty<ReaderChunk>()
        }, replaceExisting: false);
        using (DocumentReaderEngine.UseHandlerRegistry(registry.CaptureSnapshot())) {
            var options = new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent };
            Assert.Equal(packageLimit, DocumentReaderEngine.ResolveInitialMaxInputBytes(
                "large.pages", options));
            Assert.Equal(packageLimit, DocumentReaderEngine.ResolveStreamMaxInputBytes(
                "large.pages", options, streamCanSeek: false));
            Assert.Equal(packageLimit, DocumentReaderEngine.ResolveStreamMaxInputBytes(
                "large.pages", options, streamCanSeek: true));
            Assert.Equal(64L * 1024L * 1024L, DocumentReaderEngine.ResolveStreamMaxInputBytes(
                "unknown.bin", options, streamCanSeek: true));
        }
    }

    [Fact]
    public void PreferContentClampsExplicitInputBudgetToKnownHandlerCeiling() {
        var registry = new ReaderHandlerRegistry();
        registry.Register(new ReaderHandlerRegistration {
            Id = "officeimo.tests.small-iwork-ceiling",
            Kind = ReaderInputKind.IWork,
            Extensions = new[] { ".pages" },
            MaxInputBytesCeiling = 8,
            ReadPath = (_, _, _) => Array.Empty<ReaderChunk>(),
            ReadStream = (_, _, _, _) => Array.Empty<ReaderChunk>()
        }, replaceExisting: false);
        using (DocumentReaderEngine.UseHandlerRegistry(registry.CaptureSnapshot())) {
            var options = new ReaderOptions {
                DetectionMode = ReaderDetectionMode.PreferContent,
                MaxInputBytes = 1024
            };
            Assert.Equal(8, DocumentReaderEngine.ResolveInitialMaxInputBytes("input.pages", options));
            Assert.Equal(8, DocumentReaderEngine.ResolveStreamMaxInputBytes(
                "input.pages", options, streamCanSeek: false));
            Assert.Equal(1024, DocumentReaderEngine.ResolveStreamMaxInputBytes(
                "input.bin", options, streamCanSeek: false));
        }
    }

    [Fact]
    public void ChunkOnlyReadCarriesSourceWarnings() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        ReaderChunk[] chunks = reader.Read(Fixture("picodocs/sample-v14.4.pages")).ToArray();
        Assert.Contains(chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("IWORK_PAGES_TEXT_UNSUPPORTED", StringComparison.Ordinal));
    }

    [Fact]
    public void TableRowBudgetCountsDataRowsAfterHeaders() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("keynotekit/tabledeck-v15.2.1.key"),
            new ReaderOptions { MaxTableRows = 1 });
        ReaderTable table = document.Tables[0];
        Assert.Single(table.Rows);
        Assert.Equal("Product", table.Columns[0]);
        Assert.Equal(2, table.TotalRowCount);
        Assert.True(table.Truncated);
    }

    [Fact]
    public void SplitTextKeepsOneLogicalMarkdownBlock() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("picodocs/sample-v14.4.pages"), new ReaderOptions { MaxChars = 256 });

        Assert.Contains(document.Chunks, chunk => chunk.ContinuesPreviousChunk);
        string markdown = Assert.IsType<string>(document.Markdown);
        Assert.Contains("Preserve reading order", markdown, StringComparison.Ordinal);
        Assert.Equal(1, markdown.Split(new[] { "Preserve reading order" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void SplitTableMarkdownRespectsChunkBudgetAndReassembles() {
        const int maxChars = 256;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult document = reader.ReadDocument(
            Fixture("picodocs/sample-v14.4.pages"), new ReaderOptions { MaxChars = maxChars });

        ReaderChunk first = Assert.Single(document.Chunks,
            chunk => chunk.Tables?.Count > 0 && chunk.Tables[0] == document.Tables[0]);
        string? anchor = first.Location.BlockAnchor;
        ReaderChunk[] parts = document.Chunks.Where(chunk => chunk.Location.BlockAnchor == anchor).ToArray();
        Assert.True(parts.Length > 1);
        Assert.All(parts, part => Assert.InRange(part.Markdown!.Length, 0, maxChars));
        Assert.Equal(document.Tables[0].ToMarkdownTable(), string.Concat(parts.Select(part => part.Markdown)));
    }

    [Fact]
    public void IWorkResultsKeepTheirNeutralFormatIdentity() {
        Assert.Equal(OfficeDocumentFormat.IWork,
            OfficeDocumentReadResultPdfExtensions.MapFormat(ReaderInputKind.IWork));
    }

    private static string Fixture(string relativePath) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
            relativePath.Replace('/', Path.DirectorySeparatorChar));

    private static byte[] CreateZip(params string[] entryNames) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (string name in entryNames) {
                using Stream entry = archive.CreateEntry(name).Open();
                entry.WriteByte(1);
            }
        }
        return stream.ToArray();
    }

    private static byte[] CreateNestedIndexPackage(byte[] directPackage, bool preserveIndexPrefix = false) {
        using var originalStream = new MemoryStream(directPackage, writable: false);
        using var original = new ZipArchive(originalStream, ZipArchiveMode.Read);
        using var nestedStream = new MemoryStream();
        using var outerStream = new MemoryStream();
        using (var nested = new ZipArchive(nestedStream, ZipArchiveMode.Create, leaveOpen: true))
        using (var outer = new ZipArchive(outerStream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in original.Entries) {
                string name = entry.FullName;
                bool isIndex = name.StartsWith("Index/", StringComparison.Ordinal);
                string targetName = isIndex && !preserveIndexPrefix ? name.Substring("Index/".Length) : name;
                ZipArchive target = isIndex ? nested : outer;
                using Stream input = entry.Open();
                using Stream output = target.CreateEntry(targetName).Open();
                input.CopyTo(output);
            }
        }
        using (var outer = new ZipArchive(outerStream, ZipArchiveMode.Update, leaveOpen: true)) {
            using Stream index = outer.CreateEntry("Index.zip").Open();
            nestedStream.Position = 0;
            nestedStream.CopyTo(index);
        }
        return outerStream.ToArray();
    }

    private static byte[] WrapWithZip64EndRecord(byte[] source) {
        int endOffset = source.Length - 22;
        Assert.Equal(0x06054B50u, BitConverter.ToUInt32(source, endOffset));
        ushort count = BitConverter.ToUInt16(source, endOffset + 10);
        uint size = BitConverter.ToUInt32(source, endOffset + 12);
        uint offset = BitConverter.ToUInt32(source, endOffset + 16);
        using var output = new MemoryStream();
        output.Write(source, 0, endOffset);
        using (var writer = new BinaryWriter(output, System.Text.Encoding.UTF8, leaveOpen: true)) {
            writer.Write(0x06064B50u);
            writer.Write(44ul);
            writer.Write((ushort)45);
            writer.Write((ushort)45);
            writer.Write(0u);
            writer.Write(0u);
            writer.Write((ulong)count);
            writer.Write((ulong)count);
            writer.Write((ulong)size);
            writer.Write((ulong)offset);
            writer.Write(0x07064B50u);
            writer.Write(0u);
            writer.Write((ulong)endOffset);
            writer.Write(1u);
            writer.Write(0x06054B50u);
            writer.Write(0u);
            writer.Write(ushort.MaxValue);
            writer.Write(ushort.MaxValue);
            writer.Write(uint.MaxValue);
            writer.Write(uint.MaxValue);
            writer.Write((ushort)0);
        }
        return output.ToArray();
    }

    private static byte[] WrapIndexEntryWithZip64Metadata(byte[] source) {
        int endOffset = source.Length - 22;
        Assert.Equal(0x06054B50u, BitConverter.ToUInt32(source, endOffset));
        int centralOffset = (int)BitConverter.ToUInt32(source, endOffset + 16);
        int centralSize = (int)BitConverter.ToUInt32(source, endOffset + 12);
        int entryOffset = centralOffset;
        while (entryOffset < centralOffset + centralSize) {
            Assert.Equal(0x02014B50u, BitConverter.ToUInt32(source, entryOffset));
            int nameLength = BitConverter.ToUInt16(source, entryOffset + 28);
            int extraLength = BitConverter.ToUInt16(source, entryOffset + 30);
            int commentLength = BitConverter.ToUInt16(source, entryOffset + 32);
            string name = System.Text.Encoding.UTF8.GetString(source, entryOffset + 46, nameLength);
            if (name == "Index.zip") {
                const int zip64ExtraLength = 28;
                int insertion = entryOffset + 46 + nameLength;
                var result = new byte[source.Length + zip64ExtraLength];
                Array.Copy(source, 0, result, 0, insertion);
                Array.Copy(source, insertion, result, insertion + zip64ExtraLength,
                    source.Length - insertion);
                ulong expanded = BitConverter.ToUInt32(source, entryOffset + 24);
                ulong compressed = BitConverter.ToUInt32(source, entryOffset + 20);
                ulong offset = BitConverter.ToUInt32(source, entryOffset + 42);
                using (var writer = new BinaryWriter(new MemoryStream(result, writable: true))) {
                    writer.BaseStream.Position = insertion;
                    writer.Write((ushort)1);
                    writer.Write((ushort)24);
                    writer.Write(expanded);
                    writer.Write(compressed);
                    writer.Write(offset);
                }
                Array.Copy(BitConverter.GetBytes((ushort)45), 0, result, entryOffset + 6, 2);
                Array.Copy(BitConverter.GetBytes(uint.MaxValue), 0, result, entryOffset + 20, 4);
                Array.Copy(BitConverter.GetBytes(uint.MaxValue), 0, result, entryOffset + 24, 4);
                Array.Copy(BitConverter.GetBytes((ushort)(extraLength + zip64ExtraLength)), 0,
                    result, entryOffset + 30, 2);
                Array.Copy(BitConverter.GetBytes(uint.MaxValue), 0, result, entryOffset + 42, 4);
                Array.Copy(BitConverter.GetBytes((uint)(centralSize + zip64ExtraLength)), 0,
                    result, endOffset + zip64ExtraLength + 12, 4);
                return result;
            }
            entryOffset += 46 + nameLength + extraLength + commentLength;
        }
        throw new InvalidDataException("The generated nested package lacks Index.zip.");
    }
}
