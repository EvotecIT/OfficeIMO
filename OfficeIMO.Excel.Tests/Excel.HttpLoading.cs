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
            using var handler = new FakeWorkbookHttpMessageHandler((request, _) => {
                capturedRequest = request;
                return Task.FromResult(CreateWorkbookResponse(workbookBytes));
            });

            var httpOptions = new ExcelHttpLoadOptions {
                SchemePolicy = ExcelUriSchemePolicy.HttpAndHttps,
                HttpMessageHandler = handler,
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
        public async Task ExcelHttpDocumentLoadUsesTheSameReadWriteDefaultAsOtherSources() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(CreateWorkbookResponse(workbookBytes)));

            await using var document = await ExcelDocument.LoadAsync(
                new Uri("https://example.test/workbook.xlsx"),
                new ExcelHttpLoadOptions { HttpMessageHandler = handler });

            Assert.Equal(OfficeIMO.Drawing.DocumentAccessMode.ReadWrite, document.AccessMode);
            Assert.Equal("Remote", document.Sheets[0].Name);
            Assert.Null(document.FilePath);
        }

        [Fact]
        public async Task ExcelHttpDocumentLoadAsyncRejectsSaveOnDisposeBeforeDownloading() {
            int requestCount = 0;
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                requestCount++;
                return Task.FromResult(CreateWorkbookResponse(CreateRemoteWorkbookBytes()));
            });

            ArgumentException exception = await Assert.ThrowsAsync<ArgumentException>(() => ExcelDocument.LoadAsync(
                new Uri("https://example.test/workbook.xlsx"),
                new ExcelHttpLoadOptions { HttpMessageHandler = handler },
                new ExcelLoadOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose }));

            Assert.Equal("options", exception.ParamName);
            Assert.Contains("detached", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, requestCount);
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsContentLengthOverLimitBeforeCopying() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(CreateWorkbookResponse(workbookBytes)));

            var ex = await Assert.ThrowsAsync<IOException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions {
                        HttpMessageHandler = handler,
                        MaxBytes = workbookBytes.Length - 1
                    }));

            Assert.Contains("too large", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task ExcelHttpReaderHonorsTheStricterReadInputLimitBeforeCopying() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(CreateWorkbookResponse(workbookBytes)));

            var ex = await Assert.ThrowsAsync<IOException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    new ExcelReadOptions { MaxInputBytes = workbookBytes.Length - 1 },
                    new ExcelHttpLoadOptions {
                        HttpMessageHandler = handler,
                        MaxBytes = workbookBytes.Length * 2L
                    }));

            Assert.Contains((workbookBytes.Length - 1).ToString(), ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsResponseThatExceedsLimitDuringCopy() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StreamContent(new NonSeekableMemoryStream(workbookBytes))
                };
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                return Task.FromResult(response);
            });

            var ex = await Assert.ThrowsAsync<IOException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions {
                        HttpMessageHandler = handler,
                        MaxBytes = workbookBytes.Length - 1
                    }));

            Assert.Contains("exceeded", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task ExcelHttpLoadReportsProgress() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            var progress = new CapturingProgress();
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(CreateWorkbookResponse(workbookBytes)));

            using var reader = await ExcelDocumentReader.OpenAsync(
                new Uri("https://example.test/workbook.xlsx"),
                httpOptions: new ExcelHttpLoadOptions {
                    HttpMessageHandler = handler,
                    Progress = progress
                });

            Assert.Equal(new[] { "Remote" }, reader.GetSheetNames());
            var last = Assert.Single(progress.Events);
            Assert.Equal(workbookBytes.Length, last.BytesRead);
            Assert.Equal(workbookBytes.Length, last.ContentLength);
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsInvalidZipHeaderByDefault() {
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(System.Text.Encoding.UTF8.GetBytes("not a workbook"))
                };
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                return Task.FromResult(response);
            });

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions { HttpMessageHandler = handler }));
        }

        [Fact]
        public async Task ExcelHttpLoadCanValidateContentTypeWhenPresent() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                var response = CreateWorkbookResponse(workbookBytes);
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/plain");
                return Task.FromResult(response);
            });

            var ex = await Assert.ThrowsAsync<InvalidDataException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions {
                        HttpMessageHandler = handler,
                        ValidateContentTypeWhenPresent = true
                    }));

            Assert.Contains("Content-Type", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task ExcelHttpLoadAllowsMacroEnabledWorkbookContentTypeByDefault() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                var response = CreateWorkbookResponse(workbookBytes);
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(
                    "application/vnd.ms-excel.sheet.macroEnabled.12");
                return Task.FromResult(response);
            });

            using var reader = await ExcelDocumentReader.OpenAsync(
                new Uri("https://example.test/workbook.xlsm"),
                httpOptions: new ExcelHttpLoadOptions {
                    HttpMessageHandler = handler,
                    ValidateContentTypeWhenPresent = true
                });

            Assert.Equal(new[] { "Remote" }, reader.GetSheetNames());
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsHttpsToHttpRedirectByDefault() {
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                var response = new HttpResponseMessage(HttpStatusCode.Redirect);
                response.Headers.Location = new Uri("http://example.test/workbook.xlsx");
                return Task.FromResult(response);
            });

            var ex = await Assert.ThrowsAsync<NotSupportedException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions { HttpMessageHandler = handler }));

            Assert.Contains("HTTPS", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsInitialHostOutsideAllowListBeforeFetch() {
            int requestCount = 0;
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                requestCount++;
                return Task.FromResult(CreateWorkbookResponse(CreateRemoteWorkbookBytes()));
            });

            var options = new ExcelHttpLoadOptions {
                HttpMessageHandler = handler
            };
            options.AllowedHosts.Add("allowed.example.test");

            var ex = await Assert.ThrowsAsync<NotSupportedException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://blocked.example.test/workbook.xlsx"),
                    httpOptions: options));

            Assert.Contains("AllowedHosts", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, requestCount);
        }

        [Fact]
        public async Task ExcelHttpLoadRejectsRedirectHostOutsideAllowListBeforeFetch() {
            var requests = new List<HttpRequestMessage>();
            using var handler = new FakeWorkbookHttpMessageHandler((request, _) => {
                requests.Add(CloneRequestHeaders(request));
                if (requests.Count == 1) {
                    var redirect = new HttpResponseMessage(HttpStatusCode.Redirect);
                    redirect.Headers.Location = new Uri("https://blocked.example.test/workbook.xlsx");
                    return Task.FromResult(redirect);
                }

                return Task.FromResult(CreateWorkbookResponse(CreateRemoteWorkbookBytes()));
            });

            var options = new ExcelHttpLoadOptions {
                HttpMessageHandler = handler
            };
            options.AllowedHosts.Add("example.test");

            var ex = await Assert.ThrowsAsync<NotSupportedException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: options));

            Assert.Contains("AllowedHosts", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Single(requests);
            Assert.Equal("example.test", requests[0].RequestUri!.Host);
        }

        [Fact]
        public async Task ExcelHttpLoadAllowsRedirectHostInsideAllowList() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            var requests = new List<HttpRequestMessage>();
            using var handler = new FakeWorkbookHttpMessageHandler((request, _) => {
                requests.Add(CloneRequestHeaders(request));
                if (requests.Count == 1) {
                    var redirect = new HttpResponseMessage(HttpStatusCode.Redirect);
                    redirect.Headers.Location = new Uri("https://cdn.example.test/workbook.xlsx");
                    return Task.FromResult(redirect);
                }

                return Task.FromResult(CreateWorkbookResponse(workbookBytes));
            });

            var options = new ExcelHttpLoadOptions {
                HttpMessageHandler = handler
            };
            options.AllowedHosts.Add("example.test");
            options.AllowedHosts.Add("cdn.example.test");

            using var reader = await ExcelDocumentReader.OpenAsync(
                new Uri("https://example.test/workbook.xlsx"),
                httpOptions: options);

            Assert.Equal(new[] { "Remote" }, reader.GetSheetNames());
            Assert.Equal(2, requests.Count);
            Assert.Equal("cdn.example.test", requests[1].RequestUri!.Host);
        }

        [Fact]
        public async Task ExcelHttpLoadDoesNotForwardCustomHeadersAcrossRedirectedHosts() {
            byte[] workbookBytes = CreateRemoteWorkbookBytes();
            var requests = new List<HttpRequestMessage>();
            using var handler = new FakeWorkbookHttpMessageHandler((request, _) => {
                requests.Add(CloneRequestHeaders(request));
                if (requests.Count == 1) {
                    var redirect = new HttpResponseMessage(HttpStatusCode.Redirect);
                    redirect.Headers.Location = new Uri("https://cdn.example.test/workbook.xlsx");
                    return Task.FromResult(redirect);
                }

                return Task.FromResult(CreateWorkbookResponse(workbookBytes));
            });

            var options = new ExcelHttpLoadOptions {
                HttpMessageHandler = handler,
                UserAgent = "OfficeIMO.Tests"
            };
            options.Headers["X-Api-Key"] = "secret";

            using var reader = await ExcelDocumentReader.OpenAsync(
                new Uri("https://example.test/workbook.xlsx"),
                httpOptions: options);

            Assert.Equal(new[] { "Remote" }, reader.GetSheetNames());
            Assert.Equal(2, requests.Count);
            Assert.True(requests[0].Headers.Contains("X-Api-Key"));
            Assert.False(requests[1].Headers.Contains("X-Api-Key"));
            Assert.Contains(requests[1].Headers.UserAgent, value => value.Product?.Name == "OfficeIMO.Tests");
        }

        [Fact]
        public async Task ExcelHttpLoadObservesCancellationToken() {
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StreamContent(new BlockingReadStream())
                }));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAsync<TaskCanceledException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("https://example.test/workbook.xlsx"),
                    httpOptions: new ExcelHttpLoadOptions { HttpMessageHandler = handler },
                    cancellationToken: cancellation.Token));
        }

        private static byte[] CreateRemoteWorkbookBytes() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                var sheet = document.AddWorksheet("Remote");
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

        private static HttpRequestMessage CloneRequestHeaders(HttpRequestMessage request) {
            var clone = new HttpRequestMessage(request.Method, request.RequestUri);
            foreach (var header in request.Headers) {
                clone.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }

            return clone;
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
