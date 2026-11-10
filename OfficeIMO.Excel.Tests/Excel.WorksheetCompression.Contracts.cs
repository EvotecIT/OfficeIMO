using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void WorksheetCompression_PropagatesWorkerFailureWithoutReplay() {
            using var destination = new PartiallyFailingXlsbRecordStream(rejectSecondWrite: true);
            using var pipeline = new WorksheetCompressionPipeline(destination);
            byte[] payload = new byte[4 * WorksheetCompressionPipeline.ChunkSize];
            IOException error = Assert.Throws<IOException>(() => {
                pipeline.Write(payload, 0, payload.Length);
                pipeline.Complete();
            });
            Assert.Same(destination.InitialFailure, error);
            pipeline.Dispose();
            Assert.Equal(1, destination.WriteAttempts);
            Assert.Equal(8, destination.Length);
        }

        [Fact]
        public async Task WorksheetCompression_AbortKeepsProducerFailureAndJoinsConsumer() {
            using var destination = new GatedCompressionFailureStream();
            using var producerFailed = new ManualResetEventSlim();
            var initial = new InvalidOperationException("Producer validation failure.");
            Task release = Task.Run(() => {
                Assert.True(producerFailed.Wait(TimeSpan.FromSeconds(10)));
                destination.Release.Set();
            });
            try {
                Exception error = Assert.Throws<InvalidOperationException>((Action)(() => {
                    using var pipeline = new WorksheetCompressionPipeline(destination);
                    byte[] payload = new byte[WorksheetCompressionPipeline.ChunkSize];
                    pipeline.Write(payload, 0, payload.Length);
                    Assert.True(destination.Entered.Wait(TimeSpan.FromSeconds(10)));
                    producerFailed.Set();
                    throw initial;
                }));
                Assert.Same(initial, error);
                Assert.True(destination.Finished);
                Assert.Equal(1, destination.WriteAttempts);
            } finally {
                producerFailed.Set();
                destination.Release.Set();
                await release;
            }
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsb, 4095, ExcelExecutionMode.Automatic, null, false)]
        [InlineData(ExcelFileFormat.Xlsb, 4096, ExcelExecutionMode.Automatic, null, false)]
        [InlineData(ExcelFileFormat.Xlsb, 4096, ExcelExecutionMode.Sequential, null, false)]
        [InlineData(ExcelFileFormat.Xlsb, 4096, ExcelExecutionMode.Automatic, 1, false)]
        [InlineData(ExcelFileFormat.Xlsb, 4096, ExcelExecutionMode.Automatic, null, true)]
        [InlineData(ExcelFileFormat.Xlsx, 16384, ExcelExecutionMode.Automatic, null, false)]
        [InlineData(ExcelFileFormat.Xlsx, 16384, ExcelExecutionMode.Sequential, null, false)]
        [InlineData(ExcelFileFormat.Xlsx, 16384, ExcelExecutionMode.Automatic, 1, false)]
        [InlineData(ExcelFileFormat.Xlsx, 16384, ExcelExecutionMode.Automatic, null, true)]
        public void BufferedSave_PreservesCapturedSelectorsAndValuesAcrossPolicies(
            ExcelFileFormat format, int rowCount, ExcelExecutionMode mode, int? degree, bool diagnostics) {
            int caller = Thread.CurrentThread.ManagedThreadId;
            int selectorCalls = 0;
            using ExcelDocument document = ExcelDocument.Create();
            document.Execution.Mode = mode;
            document.Execution.MaxDegreeOfParallelism = degree;
            if (diagnostics) document.Execution.OnTiming = (_, _) => Assert.Equal(caller, Thread.CurrentThread.ManagedThreadId);
            var input = Enumerable.Range(0, rowCount).ToArray();
            object Observe(int row, int column, object value) {
                Assert.Equal(caller, Thread.CurrentThread.ManagedThreadId);
                Assert.Equal(row, selectorCalls / 4);
                Assert.Equal(column, selectorCalls % 4);
                selectorCalls++;
                return value;
            }
            document.AddSheet("Data").WriteData(input,
                ("Name", row => Observe(row, 0, "_x0041_ <&> " + row)),
                ("Id", row => Observe(row, 1, row)),
                ("Date", row => Observe(row, 2, new DateTime(2024, 1, 1).AddDays(row % 365))),
                ("Amount", row => Observe(row, 3, row * 1.25)));
            Assert.Equal(rowCount * 4, selectorCalls);
            byte[] bytes;
            if (format == ExcelFileFormat.Xlsb) {
                bytes = document.ToBytes(format);
            } else {
                using var destination = new BufferedSaveWriteTrackingStream();
                document.Save(destination, new ExcelSaveOptions { Format = format });
                int[] writingThreads = destination.WritingThreads;
                Assert.NotEmpty(writingThreads);
                if (mode == ExcelExecutionMode.Sequential || degree == 1 || diagnostics) {
                    Assert.All(writingThreads, thread => Assert.Equal(caller, thread));
                }
                bytes = destination.ToArray();
            }
            Assert.Equal(rowCount * 4, selectorCalls);
            Assert.Equal(format == ExcelFileFormat.Xlsb
                ? ExcelSavePackageWriter.NativeBinaryDirectPackage
                : ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            using var reader = ExcelDocument.OpenReader(bytes);
            Assert.Equal(4, reader.FieldCount);
            for (int row = 0; row < rowCount; row++) {
                Assert.True(reader.Read());
                Assert.Equal("_x0041_ <&> " + row, reader.GetString(0));
                Assert.Equal(row, reader.GetInt32(1));
                Assert.Equal(new DateTime(2024, 1, 1).AddDays(row % 365), reader.GetDateTime(2));
                Assert.Equal(row * 1.25, reader.GetDouble(3));
            }
            Assert.False(reader.Read());
            Assert.False(reader.NextResult());
        }

        private sealed class BufferedSaveWriteTrackingStream : MemoryStream {
            private readonly HashSet<int> _writingThreads = new();
            internal int[] WritingThreads {
                get { lock (_writingThreads) return _writingThreads.ToArray(); }
            }
            private void RecordWrite() {
                lock (_writingThreads) _writingThreads.Add(Thread.CurrentThread.ManagedThreadId);
            }
            public override void Write(byte[] buffer, int offset, int count) {
                RecordWrite();
                base.Write(buffer, offset, count);
            }
#if NET8_0_OR_GREATER
            public override void Write(ReadOnlySpan<byte> buffer) {
                RecordWrite();
                base.Write(buffer);
            }
#endif
            public override void WriteByte(byte value) {
                RecordWrite();
                base.WriteByte(value);
            }
        }

        private sealed class GatedCompressionFailureStream : MemoryStream {
            internal readonly ManualResetEventSlim Entered = new();
            internal readonly ManualResetEventSlim Release = new();
            internal int WriteAttempts;
            internal bool Finished;
            public override void Write(byte[] buffer, int offset, int count) {
                WriteAttempts++;
                Entered.Set();
                try {
                    if (!Release.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("Consumer gate was not released.");
                    throw new IOException("Consumer write failure.");
                } finally { Finished = true; }
            }
            protected override void Dispose(bool disposing) {
                if (disposing) { Entered.Dispose(); Release.Dispose(); }
                base.Dispose(disposing);
            }
        }
    }
}
