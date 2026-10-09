using DBAClientX.Dbf;
using OfficeIMO.Reader.Dbf;

namespace OfficeIMO.Reader.Dbf.Tests {
    public sealed class DbfReaderTests {
        internal static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

        [Fact]
        public void ReaderProjectsNativeSchemaAndValuesWithPhysicalLocations() {
            var reader = new OfficeDocumentReaderBuilder().AddDbfHandler(new ReaderDbfOptions { ChunkRows = 1 }).Build();
            ReaderChunk[] chunks = reader.Read(Fixture("plain.dbf")).ToArray();
            Assert.Equal(2, chunks.Length);
            Assert.Equal(ReaderInputKind.Dbf, chunks[0].Kind);
            ReaderTable table = Assert.Single(chunks[0].Tables!);
            Assert.Equal(new[] { "NAME", "AMOUNT", "ACTIVE" }, table.Columns);
            Assert.Equal("Café | <script> & **bold**", table.Rows[0][0]);
            Assert.Equal("1234.50", table.Rows[0][1]);
            Assert.Equal("true", table.Rows[0][2]);
            Assert.Equal(0, chunks[0].Location.SourceBlockIndex);
            Assert.Equal(2, chunks[1].Location.SourceBlockIndex);
            Assert.Contains("\\|", chunks[0].Markdown);
            Assert.Contains("\\<script\\>", chunks[0].Markdown);
            Assert.Contains("\\*\\*bold\\*\\*", chunks[0].Markdown);
            Assert.Null(chunks[0].SourceHash);
            Assert.NotNull(chunks[0].ChunkHash);
            Assert.Contains(chunks[1].Warnings!, value => value.Contains("null values"));
            var capability = Assert.Single(reader.GetCapabilities());
            Assert.True(capability.SupportsIncrementalPath);
            Assert.True(capability.SupportsIncrementalStream);
            Assert.Equal(ReaderFormatSupport.ReadConvert, Assert.Single(capability.FormatQualifications).Support);
        }

        [Fact]
        public void RegistrationCopiesSettingsAndHonorsNativeAndCoreLimits() {
            var options = new ReaderDbfOptions { ChunkRows = 20, ReadOptions = new DbfReadOptions { IncludeDeletedRecords = true } };
            var reader = new OfficeDocumentReaderBuilder().AddDbfHandler(options).Build();
            options.ReadOptions.IncludeDeletedRecords = false;
            options.ChunkRows = 1;
            ReaderChunk[] chunks = reader.Read(Fixture("plain.dbf"), new ReaderOptions { MaxTableRows = 2 }).ToArray();
            Assert.Equal(2, chunks.Length);
            Assert.Equal(2, Assert.Single(chunks[0].Tables!).Rows.Count);
            Assert.Contains(chunks[0].Warnings!, value => value.Contains("marked deleted"));
            Assert.DoesNotContain(chunks[1].Warnings!, value => value.Contains("marked deleted"));
            Assert.ThrowsAny<IOException>(() => reader.Read(Fixture("plain.dbf"), new ReaderOptions { MaxInputBytes = 40 }).ToArray());
        }

        [Fact]
        public void SidecarReadsAreExplicitAndCannotBeInferredFromAStreamName() {
            var safe = new OfficeDocumentReaderBuilder().AddDbfHandler().Build();
            Assert.Throws<InvalidDataException>(() => safe.Read(Fixture("db3.dbf")).ToArray());
            var enabled = new OfficeDocumentReaderBuilder().AddDbfHandler(new ReaderDbfOptions { AllowMemoSidecarReads = true }).Build();
            ReaderChunk chunk = Assert.Single(enabled.Read(Fixture("vfp.dbf")));
            Assert.Equal("VFP memo", Assert.Single(chunk.Tables!).Rows[0][5]);
            Assert.Equal("AP8B", Assert.Single(chunk.Tables!).Rows[0][7]);
            Assert.Contains(chunk.Warnings!, value => value.Contains("Base64"));
            using var stream = File.OpenRead(Fixture("db3.dbf"));
            Assert.Throws<InvalidDataException>(() => enabled.Read(stream, Fixture("db3.dbf")).ToArray());
            Assert.True(stream.CanRead);
        }

        [Fact]
        public void ContentFirstProbeValidatesNativeSchemaWithoutClosingCallerStream() {
            var reader = new OfficeDocumentReaderBuilder().AddDbfHandler().Build();
            using var stream = File.OpenRead(Fixture("plain.dbf"));
            stream.Position = 8;
            ReaderChunk chunk = Assert.Single(reader.Read(stream, "table.dbf", new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent }));
            Assert.Equal(ReaderInputKind.Dbf, chunk.Kind);
            Assert.True(stream.CanRead);
            Assert.Equal(8, stream.Position);
        }

        [Fact]
        public void EarlyIncrementalTerminationLeavesCallerOpenAndObservesCancellation() {
            var reader = new OfficeDocumentReaderBuilder().AddDbfHandler(new ReaderDbfOptions { ChunkRows = 1 }).Build();
            using var stream = File.OpenRead(Fixture("plain.dbf"));
            using var cancellation = new CancellationTokenSource();
            using var chunks = reader.EnumerateChunks(stream, "plain.dbf", new ReaderOptions { ComputeHashes = false }, cancellation.Token).GetEnumerator();
            Assert.True(chunks.MoveNext());
            cancellation.Cancel();
            Assert.ThrowsAny<OperationCanceledException>(() => chunks.MoveNext());
            chunks.Dispose();
            Assert.True(stream.CanRead);
        }
    }
}
