using OfficeIMO.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task ExcelHttpLoadRejectsHttpByDefault() {
            var ex = await Assert.ThrowsAsync<NotSupportedException>(() =>
                ExcelDocumentReader.OpenAsync(new Uri("http://example.test/workbook.xlsx")));

            Assert.Contains("HTTPS", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task ExcelHttpReaderAllowsHttpWhenOptedInAndSendsHeaders() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            HttpRequestMessage? capturedRequest = null;
            using var httpClient = new HttpClient(new FakeWorkbookHttpMessageHandler((request, _) => {
                capturedRequest = request;
                return Task.FromResult(CreateWorkbookResponse(workbookBytes));
            }));

            var httpOptions = new ExcelHttpLoadOptions {
                SchemePolicy = ExcelUriSchemePolicy.HttpAndHttps,
                HttpClient = httpClient,
                UserAgent = "OfficeIMO.Tests"
            };
            httpOptions.Headers["X-Test"] = "remote-load";

            using var reader = await ExcelDocumentReader.OpenAsync(
                new Uri("http://example.test/workbook.xlsx"),
                httpOptions: httpOptions);

            Assert.Equal(new[] { "Remote" }, reader.GetSheetNames());
            Assert.NotNull(capturedRequest);
            Assert.True(capturedRequest!.Headers.TryGetValues("X-Test", out var values));
            Assert.Contains("remote-load", values);
            Assert.Contains(capturedRequest.Headers.UserAgent, value => value.Product?.Name == "OfficeIMO.Tests");
        }

        [Fact]
        public async Task ExcelHttpDocumentLoadOpensDownloadedWorkbookReadOnlyByDefault() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var httpClient = new HttpClient(new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(CreateWorkbookResponse(workbookBytes))));

            await using var document = await ExcelDocument.LoadAsync(
                new Uri("https://example.test/workbook.xlsx"),
                new ExcelHttpLoadOptions { HttpClient = httpClient });

            Assert.Equal(FileAccess.Read, document.FileOpenAccess);
            Assert.Equal("Remote", document.Sheets[0].Name);
            Assert.Equal(string.Empty, document.FilePath);
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsContentLengthOverLimitBeforeCopying() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var httpClient = new HttpClient(new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(CreateWorkbookResponse(workbookBytes))));

            var ex = await Assert.ThrowsAsync<IOException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions {
                        HttpClient = httpClient,
                        MaxBytes = workbookBytes.Length - 1
                    }));

            Assert.Contains("too large", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsResponseThatExceedsLimitDuringCopy() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var httpClient = new HttpClient(new FakeWorkbookHttpMessageHandler((_, _) => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StreamContent(new NonSeekableMemoryStream(workbookBytes))
                };
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                return Task.FromResult(response);
            }));

            var ex = await Assert.ThrowsAsync<IOException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions {
                        HttpClient = httpClient,
                        MaxBytes = workbookBytes.Length - 1
                    }));

            Assert.Contains("exceeded", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task ExcelHttpLoadReportsProgress() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            var progress = new CapturingProgress();
            using var httpClient = new HttpClient(new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(CreateWorkbookResponse(workbookBytes))));

            using var reader = await ExcelDocumentReader.OpenAsync(
                new Uri("https://example.test/workbook.xlsx"),
                httpOptions: new ExcelHttpLoadOptions {
                    HttpClient = httpClient,
                    Progress = progress
                });

            Assert.Equal(new[] { "Remote" }, reader.GetSheetNames());
            var last = Assert.Single(progress.Events);
            Assert.Equal(workbookBytes.Length, last.BytesRead);
            Assert.Equal(workbookBytes.Length, last.ContentLength);
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsInvalidZipHeaderByDefault() {
            using var httpClient = new HttpClient(new FakeWorkbookHttpMessageHandler((_, _) => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(System.Text.Encoding.UTF8.GetBytes("not a workbook"))
                };
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                return Task.FromResult(response);
            }));

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions { HttpClient = httpClient }));
        }

        [Fact]
        public async Task ExcelHttpLoadCanValidateContentTypeWhenPresent() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var httpClient = new HttpClient(new FakeWorkbookHttpMessageHandler((_, _) => {
                var response = CreateWorkbookResponse(workbookBytes);
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/plain");
                return Task.FromResult(response);
            }));

            var ex = await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions {
                        HttpClient = httpClient,
                        ValidateContentTypeWhenPresent = true
                    }));

            Assert.Contains("Content-Type", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task ExcelHttpLoadObservesCancellationToken() {
            using var httpClient = new HttpClient(new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StreamContent(new BlockingReadStream())
                })));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAsync<TaskCanceledException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions { HttpClient = httpClient },
                    cancellationToken: cancellation.Token));
        }

        private static byte[] CreateRemoteWorkbookBytes() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("Remote");
                sheet.CellValue(1, 1, "Value");
            }

            return memory.ToArray();
        }

        private static HttpResponseMessage CreateWorkbookResponse(byte[] workbookBytes) {
            var response = new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(workbookBytes)
            };
            response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            return response;
        }

        private sealed class CapturingProgress : IProgress<ExcelHttpLoadProgress> {
            internal List<ExcelHttpLoadProgress> Events { get; } = new List<ExcelHttpLoadProgress>();

            public void Report(ExcelHttpLoadProgress value) {
                Events.Add(value);
            }
        }

        private sealed class FakeWorkbookHttpMessageHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _handler;

            internal FakeWorkbookHttpMessageHandler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> handler) {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request, cancellationToken);
            }
        }

        private sealed class NonSeekableMemoryStream : Stream {
            private readonly MemoryStream _inner;

            internal NonSeekableMemoryStream(byte[] bytes) {
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

            public override void Flush() {
            }

            public override int Read(byte[] buffer, int offset, int count) {
                return _inner.Read(buffer, offset, count);
            }

            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
                return _inner.ReadAsync(buffer, offset, count, cancellationToken);
            }

            public override long Seek(long offset, SeekOrigin origin) {
                throw new NotSupportedException();
            }

            public override void SetLength(long value) {
                throw new NotSupportedException();
            }

            public override void Write(byte[] buffer, int offset, int count) {
                throw new NotSupportedException();
            }

            protected override void Dispose(bool disposing) {
                if (disposing) {
                    _inner.Dispose();
                }

                base.Dispose(disposing);
            }
        }

        private sealed class BlockingReadStream : Stream {
            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush() {
            }

            public override int Read(byte[] buffer, int offset, int count) {
                Thread.Sleep(Timeout.Infinite);
                return 0;
            }

            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
                return Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken).ContinueWith(
                    static task => {
                        task.GetAwaiter().GetResult();
                        return 0;
                    },
                    CancellationToken.None,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }

            public override long Seek(long offset, SeekOrigin origin) {
                throw new NotSupportedException();
            }

            public override void SetLength(long value) {
                throw new NotSupportedException();
            }

            public override void Write(byte[] buffer, int offset, int count) {
                throw new NotSupportedException();
            }
        }
    }
}
