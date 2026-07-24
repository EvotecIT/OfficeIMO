using OfficeIMO.Excel;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task LoadApisEnforceInputBudgetForPathAndStreams() {
            byte[] workbook = CreateBatch12Workbook();
            var restrictive = new ExcelLoadOptions {
                MaxInputBytes = workbook.Length - 1L
            };
            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                File.WriteAllBytes(path, workbook);
                Assert.Throws<InvalidDataException>(() =>
                    ExcelDocument.Load(path, restrictive));

                using var stream = new Batch12NonSeekableStream(workbook);
                Assert.Throws<InvalidDataException>(() =>
                    ExcelDocument.Load(stream, restrictive));

                using var asyncStream = new Batch12NonSeekableStream(workbook);
                await Assert.ThrowsAsync<InvalidDataException>(() =>
                    ExcelDocument.LoadAsync(asyncStream, restrictive));

                using var validStream = new MemoryStream(workbook,
                    writable: false);
                using ExcelDocument valid = ExcelDocument.Load(validStream,
                    new ExcelLoadOptions { MaxInputBytes = workbook.Length });
                Assert.Single(valid.Sheets);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task RemoteLoadEnforcesInputBudgetDuringDownload() {
            byte[] workbook = CreateRemoteWorkbookBytes();
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(CreateWorkbookResponse(workbook)));

            IOException exception = await Assert.ThrowsAsync<IOException>(() =>
                ExcelDocument.LoadAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    new ExcelHttpLoadOptions {
                        HttpMessageHandler = handler,
                        MaxBytes = workbook.Length * 2L
                    },
                    new ExcelLoadOptions {
                        MaxInputBytes = workbook.Length - 1L
                    }));

            Assert.Contains((workbook.Length - 1).ToString(),
                exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task RemoteLoadRejectsInvalidInputBudgetBeforeRequest() {
            int requestCount = 0;
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                requestCount++;
                return Task.FromResult(
                    CreateWorkbookResponse(CreateRemoteWorkbookBytes()));
            });

            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
                ExcelDocument.LoadAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    new ExcelHttpLoadOptions { HttpMessageHandler = handler },
                    new ExcelLoadOptions { MaxInputBytes = 0 }));

            Assert.Equal(0, requestCount);
        }

        private static byte[] CreateBatch12Workbook() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data").CellValue(1, 1, "safe");
            return document.ToBytes();
        }

        private sealed class Batch12NonSeekableStream : Stream {
            private readonly MemoryStream _inner;

            internal Batch12NonSeekableStream(byte[] bytes) {
                _inner = new MemoryStream(bytes, writable: false);
            }

            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) =>
                _inner.Read(buffer, offset, count);
            public override long Seek(long offset, SeekOrigin origin) =>
                throw new NotSupportedException();
            public override void SetLength(long value) =>
                throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) =>
                throw new NotSupportedException();
            protected override void Dispose(bool disposing) {
                if (disposing) _inner.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
