using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Web;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderWebTests {
    [Fact]
    public async Task WebReader_RoutesBytesThroughConfiguredReaderAndAddsSafeMetadata() {
        DateTimeOffset lastModified = new DateTimeOffset(2026, 7, 15, 12, 0, 0, TimeSpan.Zero);
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            HttpResponseMessage response = TextResponse("<h1>Remote guide</h1><p>Body</p>", "text/html");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
                FileName = "\"remote.html\""
            };
            response.Content.Headers.LastModified = lastModified;
            return Task.FromResult(response);
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();
        OfficeDocumentWebReader webReader = reader.CreateWebReader(httpClient, new ReaderWebOptions {
            AllowedHosts = new[] { "example.test" }
        });

        OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
            new Uri("https://example.test/files/download?token=secret"));

        Assert.Equal(ReaderInputKind.Html, result.Kind);
        Assert.Equal("remote.html", result.Source.Path);
        Assert.Contains("Remote guide", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("officeimo.reader.web", result.CapabilitiesUsed);
        OfficeDocumentMetadataEntry requestUri = Assert.Single(
            result.Metadata,
            item => item.Id == "reader-web-request-uri");
        Assert.DoesNotContain("token", requestUri.Value, StringComparison.Ordinal);
        Assert.Equal(lastModified.UtcDateTime, result.Source.LastWriteUtc);
        Assert.All(result.Chunks, chunk => {
            Assert.Equal(result.Source.SourceId, chunk.SourceId);
            Assert.Equal(lastModified.UtcDateTime, chunk.SourceLastWriteUtc);
        });
        Assert.Equal(1, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_DerivesSourceIdentityFromTheFinalUriInsteadOfTheFileName() {
        int responseIndex = 0;
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            int current = Interlocked.Increment(ref responseIndex);
            HttpResponseMessage response = TextResponse("same body", "text/plain");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
                FileName = "\"report.txt\""
            };
            response.RequestMessage = new HttpRequestMessage(
                HttpMethod.Get,
                "https://cdn.example/tenant-" + current.ToString(CultureInfo.InvariantCulture) + "/report.txt?token=secret");
            return Task.FromResult(response);
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        OfficeDocumentReadResult first = await webReader.ReadDocumentAsync(new Uri("https://example.test/report"));
        OfficeDocumentReadResult second = await webReader.ReadDocumentAsync(new Uri("https://example.test/report"));

        Assert.Equal("report.txt", first.Source.Path);
        Assert.Equal(first.Source.Path, second.Source.Path);
        Assert.NotEqual(first.Source.SourceId, second.Source.SourceId);
        Assert.DoesNotContain("cdn.example", first.Source.SourceId, StringComparison.Ordinal);
        Assert.All(first.Chunks, chunk => Assert.Equal(first.Source.SourceId, chunk.SourceId));
        Assert.All(second.Chunks, chunk => Assert.Equal(second.Source.SourceId, chunk.SourceId));
        Assert.NotEqual(Assert.Single(first.Chunks).ChunkHash, Assert.Single(second.Chunks).ChunkHash);
    }

    [Fact]
    public async Task WebReader_UsesTheRichPipelineForMarkdownConvenience() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("# Remote note\n\nBody", "text/markdown")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        string markdown = await webReader.ConvertToMarkdownAsync(
            new Uri("https://example.test/note.md"));

        Assert.Contains("# Remote note", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WebReader_RejectsPrivateIpLiteralsBeforeSending() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("http://127.0.0.1/private.txt")));

        Assert.Equal(0, handler.CallCount);
    }

    [Theory]
    [InlineData("192.0.0.8")]
    [InlineData("192.0.0.170")]
    [InlineData("192.0.0.171")]
    [InlineData("192.88.99.1")]
    public async Task WebReader_RejectsNonGlobalIanaProtocolAssignmentTargets(string host) {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("http://" + host + "/special.txt")));

        Assert.Equal(0, handler.CallCount);
    }

    [Theory]
    [InlineData("192.0.0.9")]
    [InlineData("192.0.0.10")]
    [InlineData("192.0.1.1")]
    public async Task WebReader_DoesNotOverblockGloballyReachableOrOrdinaryTargets(string host) {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("public fixture", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
            new Uri("http://" + host + "/public.txt"));

        Assert.Contains("public fixture", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(1, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsLocalUseNat64TargetsBeforeSending() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("http://[64:ff9b:1::a00:1]/private.txt")));

        Assert.Equal(0, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsPrivateIpv4EmbeddedInTheWellKnownNat64Prefix() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("http://[64:ff9b::a00:1]/private.txt")));

        Assert.Equal(0, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_AllowsPublicIpv4EmbeddedInTheWellKnownNat64Prefix() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("public fixture", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
            new Uri("http://[64:ff9b::808:808]/public.txt"));

        Assert.Contains("public fixture", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(1, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsPrivateIpv4EmbeddedInTheIpv4TranslatedPrefix() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("http://[::ffff:0:10.0.0.1]/private.txt")));

        Assert.Equal(0, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_AllowsPublicIpv4EmbeddedInTheIpv4TranslatedPrefix() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("public fixture", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
            new Uri("http://[::ffff:0:8.8.8.8]/public.txt"));

        Assert.Contains("public fixture", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(1, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsTheIpv6DiscardOnlyPrefix() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("http://[100::1]/discard.txt")));

        Assert.Equal(0, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_DoesNotOverblockOutsideTheIpv6DiscardOnlyPrefix() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("public fixture", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
            new Uri("http://[100:0:0:1::1]/public.txt"));

        Assert.Contains("public fixture", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(1, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsPrivateIpv4EmbeddedInA6To4Target() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("http://[2002:a00:1::]/private.txt")));

        Assert.Equal(0, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_AllowsPublicIpv4EmbeddedInA6To4Target() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("public fixture", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
            new Uri("http://[2002:808:808::]/public.txt"));

        Assert.Contains("public fixture", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(1, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsAnInvalidSourceNameBeforeSending() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ArgumentException>(() => webReader.ReadDocumentAsync(
            new Uri("https://example.test/file.txt"),
            sourceName: "bad\nname.txt"));

        Assert.Equal(0, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_CanExplicitlyAllowLocalhost() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("local fixture", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions {
                AllowedHosts = new[] { "localhost" },
                AllowLocalhostAndNonPublicIpLiterals = true
            });

        OfficeDocumentReadResult result = await webReader.ReadDocumentAsync(
            new Uri("http://localhost/fixture.txt"));

        Assert.Contains("local fixture", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(1, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsAHostOutsideTheAllowlistBeforeSending() {
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(TextResponse("not reached", "text/plain")));
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions { AllowedHosts = new[] { "allowed.example" } });

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://other.example/file.txt")));

        Assert.Equal(0, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsADisallowedFinalUriReportedByTheResponse() {
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            HttpResponseMessage response = TextResponse("not parsed", "text/plain");
            response.RequestMessage = new HttpRequestMessage(HttpMethod.Get, "http://127.0.0.1/final.txt");
            return Task.FromResult(response);
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://example.test/start.txt")));

        Assert.Equal(1, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_RejectsDeclaredContentLengthAboveTheBound() {
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            var content = new ByteArrayContent(new byte[32]);
            content.Headers.ContentLength = 32;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = content });
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions { MaxResponseBytes = 16 });

        IOException exception = await Assert.ThrowsAsync<IOException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://example.test/file.bin")));

        Assert.Contains("effective web input byte limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task WebReader_AppliesTheSelectedHandlersDefaultBeforeReadingTheBody() {
        bool readerInvoked = false;
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            var content = new ByteArrayContent(new byte[32]);
            content.Headers.ContentLength = 32;
            content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
                FileName = "\"fixture.bounded\""
            };
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = content });
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = "test.bounded",
                Extensions = new[] { ".bounded" },
                DefaultMaxInputBytes = 16,
                ReadDocumentStream = (stream, sourceName, options, cancellationToken) => {
                    readerInvoked = true;
                    return new OfficeDocumentReadResult();
                }
            })
            .Build();
        OfficeDocumentWebReader webReader = reader.CreateWebReader(
            httpClient,
            new ReaderWebOptions { MaxResponseBytes = 64 });

        IOException exception = await Assert.ThrowsAsync<IOException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://example.test/download")));

        Assert.Contains("effective web input byte limit", exception.Message, StringComparison.Ordinal);
        Assert.Contains("32 > 16", exception.Message, StringComparison.Ordinal);
        Assert.False(readerInvoked);
    }

    [Fact]
    public async Task WebReader_RejectsAStreamingBodyWhenActualBytesCrossTheBound() {
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            var content = new StreamContent(new NonSeekableReadStream(new byte[32]));
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = content });
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions { MaxResponseBytes = 16 });

        IOException exception = await Assert.ThrowsAsync<IOException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://example.test/file.bin")));

        Assert.Contains("effective web input byte limit", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null, 0)]
    [InlineData(1024L, 1024)]
    [InlineData(65536L, 65536)]
    [InlineData(1073741824L, 65536)]
    public void WebReader_CapsInitialBufferCapacityIndependentOfTheResponseLimit(
        long? declaredLength,
        int expectedCapacity) {
        Assert.Equal(expectedCapacity, ReaderWebTransport.GetInitialBufferCapacity(declaredLength));
    }

    [Fact]
    public async Task WebReader_ReportsItsOwnRequestTimeout() {
        var handler = new DelegateHttpHandler(async (request, cancellationToken) => {
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return TextResponse("not reached", "text/plain");
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions { RequestTimeout = TimeSpan.FromMilliseconds(100) });

        await Assert.ThrowsAsync<TimeoutException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://example.test/slow.txt")));
    }

    [Fact]
    public async Task WebReader_TimesOutANonCooperativeSendAndReleasesItsRequestSlot() {
        var pendingSend = new TaskCompletionSource<HttpResponseMessage>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        int responseIndex = 0;
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            if (Interlocked.Increment(ref responseIndex) == 1) {
                return pendingSend.Task;
            }
            return Task.FromResult(TextResponse("recovered", "text/plain"));
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions {
                MaxConcurrentRequests = 1,
                RequestTimeout = TimeSpan.FromMilliseconds(100)
            });

        await Assert.ThrowsAsync<TimeoutException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://example.test/stalled-send.txt")));
        OfficeDocumentReadResult recovered = await webReader.ReadDocumentAsync(
            new Uri("https://example.test/recovered.txt"));
        var lateStream = new TrackingMemoryStream(Encoding.UTF8.GetBytes("late"));
        pendingSend.TrySetResult(new HttpResponseMessage(HttpStatusCode.OK) {
            Content = new StreamContent(lateStream)
        });

        Assert.Contains("recovered", recovered.Markdown, StringComparison.Ordinal);
        Assert.True(SpinWait.SpinUntil(() => lateStream.IsDisposed, TimeSpan.FromSeconds(2)));
        Assert.Equal(2, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_TimesOutANonCooperativeBodyAndReleasesItsRequestSlot() {
        var blockingStream = new CancellationIgnoringReadStream();
        int responseIndex = 0;
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            if (Interlocked.Increment(ref responseIndex) == 1) {
                var content = new StreamContent(blockingStream);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = content });
            }
            return Task.FromResult(TextResponse("recovered", "text/plain"));
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions {
                MaxConcurrentRequests = 1,
                RequestTimeout = TimeSpan.FromMilliseconds(100)
            });

        await Assert.ThrowsAsync<TimeoutException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://example.test/stalled.txt")));
        OfficeDocumentReadResult recovered = await webReader.ReadDocumentAsync(
            new Uri("https://example.test/recovered.txt"));

        Assert.Contains("recovered", recovered.Markdown, StringComparison.Ordinal);
        Assert.True(blockingStream.IsDisposed);
        Assert.Equal(2, handler.CallCount);
    }

    [Fact]
    public async Task WebReader_TimesOutANonCooperativeStreamOpenAndDisposesALateStream() {
        var stalledContent = new CancellationIgnoringOpenContent();
        int responseIndex = 0;
        var handler = new DelegateHttpHandler((request, cancellationToken) => {
            if (Interlocked.Increment(ref responseIndex) == 1) {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = stalledContent });
            }
            return Task.FromResult(TextResponse("recovered", "text/plain"));
        });
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions {
                MaxConcurrentRequests = 1,
                RequestTimeout = TimeSpan.FromMilliseconds(100)
            });

        await Assert.ThrowsAsync<TimeoutException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://example.test/stalled-open.txt")));
        OfficeDocumentReadResult recovered = await webReader.ReadDocumentAsync(
            new Uri("https://example.test/recovered.txt"));
        var lateStream = new TrackingMemoryStream(Encoding.UTF8.GetBytes("late"));
        stalledContent.CompleteOpen(lateStream);

        Assert.Contains("recovered", recovered.Markdown, StringComparison.Ordinal);
        Assert.True(stalledContent.IsDisposed);
        Assert.True(SpinWait.SpinUntil(() => lateStream.IsDisposed, TimeSpan.FromSeconds(2)));
        Assert.Equal(2, handler.CallCount);
    }

    [Fact]
    public async Task WebTransport_HandsReaderItsPreparedSnapshotWithoutCopying() {
        byte[] payload = Encoding.UTF8.GetBytes("prepared snapshot");
        var handler = new DelegateHttpHandler((request, cancellationToken) =>
            Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(payload)
            }));
        using var httpClient = new HttpClient(handler);
        using ReaderWebDownload download = await ReaderWebTransport.DownloadAsync(
            httpClient,
            new Uri("https://example.test/prepared.txt"),
            sourceName: null,
            resolveMaxResponseBytes: _ => 1024,
            new ReaderWebOptions(),
            CancellationToken.None);

        Stream prepared = await ReaderInputLimits.EnsureSeekableReadStreamAsync(
            download.Content,
            maxInputBytes: 1024);

        Assert.Same(download.Content, prepared);
        Assert.True(ReaderInputLimits.IsSnapshotStream(prepared));
        Assert.Equal(payload.Length, prepared.Length);
        Assert.Equal(0, prepared.Position);
    }

    [Fact]
    public async Task WebReader_BoundsConcurrentOperationsPerInstance() {
        var handler = new BlockingFirstRequestHandler();
        using var httpClient = new HttpClient(handler);
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            httpClient,
            new ReaderWebOptions { MaxConcurrentRequests = 1 });

        Task<OfficeDocumentReadResult> first = webReader.ReadDocumentAsync(
            new Uri("https://example.test/first.txt"));
        Task firstArrival = await Task.WhenAny(handler.FirstRequestArrived, Task.Delay(TimeSpan.FromSeconds(5)));
        Assert.Same(handler.FirstRequestArrived, firstArrival);

        Task<OfficeDocumentReadResult> second = webReader.ReadDocumentAsync(
            new Uri("https://example.test/second.txt"));
        await Task.Delay(100);
        Assert.Equal(1, handler.CallCount);

        handler.ReleaseFirstRequest();
        await Task.WhenAll(first, second);
        Assert.Equal(2, handler.CallCount);
    }

    private static HttpResponseMessage TextResponse(string value, string mediaType) {
        return new HttpResponseMessage(HttpStatusCode.OK) {
            Content = new StringContent(value, Encoding.UTF8, mediaType)
        };
    }

    private sealed class DelegateHttpHandler : HttpMessageHandler {
        private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _send;
        private int _callCount;

        internal DelegateHttpHandler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) {
            _send = send;
        }

        internal int CallCount => Volatile.Read(ref _callCount);

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken) {
            Interlocked.Increment(ref _callCount);
            return _send(request, cancellationToken);
        }
    }

    private sealed class BlockingFirstRequestHandler : HttpMessageHandler {
        private readonly TaskCompletionSource<bool> _firstRequestArrived =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<bool> _releaseFirstRequest =
            new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _callCount;

        internal Task FirstRequestArrived => _firstRequestArrived.Task;
        internal int CallCount => Volatile.Read(ref _callCount);

        internal void ReleaseFirstRequest() => _releaseFirstRequest.TrySetResult(true);

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken) {
            int call = Interlocked.Increment(ref _callCount);
            if (call == 1) {
                _firstRequestArrived.TrySetResult(true);
                Task cancellation = Task.Delay(Timeout.Infinite, cancellationToken);
                Task completed = await Task.WhenAny(_releaseFirstRequest.Task, cancellation);
                cancellationToken.ThrowIfCancellationRequested();
                await completed;
            }
            return TextResponse("request " + call.ToString(CultureInfo.InvariantCulture), "text/plain");
        }
    }

    private sealed class NonSeekableReadStream : Stream {
        private readonly MemoryStream _inner;

        internal NonSeekableReadStream(byte[] bytes) {
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
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    private sealed class CancellationIgnoringReadStream : Stream {
        private readonly TaskCompletionSource<int> _pendingRead =
            new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _isDisposed;

        internal bool IsDisposed => Volatile.Read(ref _isDisposed) != 0;

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override Task<int> ReadAsync(
            byte[] buffer,
            int offset,
            int count,
            CancellationToken cancellationToken) => _pendingRead.Task;
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            if (disposing) Interlocked.Exchange(ref _isDisposed, 1);
            base.Dispose(disposing);
        }
    }

    private sealed class CancellationIgnoringOpenContent : HttpContent {
        private readonly TaskCompletionSource<Stream> _pendingOpen =
            new TaskCompletionSource<Stream>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _isDisposed;

        internal bool IsDisposed => Volatile.Read(ref _isDisposed) != 0;

        internal void CompleteOpen(Stream stream) => _pendingOpen.TrySetResult(stream);

        protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) {
            return Task.CompletedTask;
        }

        protected override bool TryComputeLength(out long length) {
            length = 0;
            return false;
        }

        protected override Task<Stream> CreateContentReadStreamAsync() => _pendingOpen.Task;

#if !NET472
        protected override Task<Stream> CreateContentReadStreamAsync(CancellationToken cancellationToken) =>
            _pendingOpen.Task;
#endif

        protected override void Dispose(bool disposing) {
            if (disposing) Interlocked.Exchange(ref _isDisposed, 1);
            base.Dispose(disposing);
        }
    }

    private sealed class TrackingMemoryStream : MemoryStream {
        private int _isDisposed;

        internal TrackingMemoryStream(byte[] bytes) : base(bytes, writable: false) {
        }

        internal bool IsDisposed => Volatile.Read(ref _isDisposed) != 0;

        protected override void Dispose(bool disposing) {
            if (disposing) Interlocked.Exchange(ref _isDisposed, 1);
            base.Dispose(disposing);
        }
    }
}
