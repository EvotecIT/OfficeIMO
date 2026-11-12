using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

namespace OfficeIMO.Workflows.Tests;

public sealed class IppPrintClientTests {
    private static readonly Uri Printer = new("ipp://printer.test/ipp/print");

    [Fact]
    public async Task MapsIppToHttpOnPort631AndStreamsTheDocumentAfterTheAttributes() {
        IppMessage? received = null;
        byte[]? document = null;
        Uri? target = null;
        var handler = FakeHandler.Ipp(async request => {
            target = request.RequestUri;
            Assert.Equal("application/ipp", request.Content!.Headers.ContentType!.MediaType);
            byte[] body = await request.Content.ReadAsByteArrayAsync();
            received = IppCodec.Decode(body, body.Length, null, out int offset);
            document = body.Skip(offset).ToArray();
            return Reply(received, 0x0000, job => job.Add("job-id", IppValue.OfInteger(12)).Add("job-state", IppValue.OfEnum(3)));
        });
        using var client = new IppPrintClient(Printer, handler, new IppPrintClientOptions { RequestingUserName = "przemek" });

        IppSubmission submission = await client.PrintAsync(new byte[] { 1, 2, 3 }, "application/pdf", new IppJobOptions {
            JobName = "Quarterly", Copies = 2, Sides = "two-sided-long-edge",
            MediaCol = new IppMedia(21000, 29700) { Source = "main" }, PageRanges = new[] { new IppPageRange(1, 2) }
        });

        Assert.Equal(new Uri("http://printer.test:631/ipp/print"), target);
        Assert.Equal(new byte[] { 1, 2, 3 }, document);
        IppAttributeGroup operation = received!.FirstGroup(IppGroupTag.Operation)!;
        Assert.Equal(new[] { "attributes-charset", "attributes-natural-language", "printer-uri" }, operation.Attributes.Take(3).Select(a => a.Name));
        Assert.Equal("ipp://printer.test/ipp/print", operation["printer-uri"]!.First!.Text);
        Assert.Equal("przemek", operation["requesting-user-name"]!.First!.Text);
        Assert.Equal("application/pdf", operation["document-format"]!.First!.Text);
        IppAttributeGroup job = received.FirstGroup(IppGroupTag.Job)!;
        Assert.Equal(2, job["copies"]!.First!.Integer);
        Assert.Equal(21000, job["media-col"]!.First!.Members!.Single(m => m.Name == "media-size").First!.Members![0].First!.Integer);
        Assert.Equal(12, submission.Job.Id);
        Assert.Equal(IppJobState.Pending, submission.Job.State);
    }

    [Fact]
    public async Task UsesChunkedTransferForStreamsOfUnknownLength() {
        bool? chunked = null;
        long? length = null;
        var handler = FakeHandler.Ipp(async request => {
            chunked = request.Headers.TransferEncodingChunked;
            length = request.Content!.Headers.ContentLength;
            byte[] body = await request.Content.ReadAsByteArrayAsync();
            IppMessage message = IppCodec.Decode(body, body.Length, null, out int offset);
            Assert.Equal(200_000, body.Length - offset);
            return Reply(message, 0, job => job.Add("job-id", IppValue.OfInteger(1)).Add("job-state", IppValue.OfEnum(5)));
        });
        using var client = new IppPrintClient(Printer, handler);

        await client.PrintAsync(new NonSeekableStream(new byte[200_000]), "image/pwg-raster");

        Assert.True(chunked);
        Assert.Null(length);
    }

    [Fact]
    public async Task ConnectionLossAfterTheDocumentStartedIsUncertainDelivery() {
        var handler = FakeHandler.Http(async (request, _) => {
            await request.Content!.ReadAsByteArrayAsync();
            throw new HttpRequestException("Connection reset by peer");
        });
        using var client = new IppPrintClient(Printer, handler);

        PdfPrintDeliveryException error = await Assert.ThrowsAsync<PdfPrintDeliveryException>(() => client.PrintAsync(new byte[16], "application/pdf"));

        Assert.Null(error.JobId);
        Assert.IsType<IppException>(error.InnerException);
    }

    [Fact]
    public async Task ConnectionFailureBeforeTheBodyIsAnOrdinaryFailure() {
        using var client = new IppPrintClient(Printer, FakeHandler.Http((_, _) => throw new HttpRequestException("Connection refused")));

        IppException error = await Assert.ThrowsAsync<IppException>(() => client.PrintAsync(new byte[16], "application/pdf"));

        Assert.Null(error.StatusCode);
    }

    [Fact]
    public async Task AcceptanceWithoutAJobIdIsUncertainDelivery() {
        var handler = FakeHandler.Ipp(async request => Reply(await Decode(request), 0, null));
        using var client = new IppPrintClient(Printer, handler);

        await Assert.ThrowsAsync<PdfPrintDeliveryException>(() => client.PrintAsync(new byte[16], "application/pdf"));
    }

    [Fact]
    public async Task ARejectionStatusIsDefinitiveAndCarriesTheStatusMessage() {
        var handler = FakeHandler.Ipp(async request => {
            IppMessage response = Reply(await Decode(request), 0x040A, null);
            response.Groups[0].Add("status-message", IppValue.OfString(IppValueTag.TextWithoutLanguage, "Unsupported format"));
            return response;
        });
        using var client = new IppPrintClient(Printer, handler);

        IppException error = await Assert.ThrowsAsync<IppException>(() => client.PrintAsync(new byte[16], "application/x-unknown"));

        Assert.Equal(0x040A, error.StatusCode);
        Assert.Contains("Unsupported format", error.Message);
    }

    [Fact]
    public async Task RefusesRedirects() {
        var handler = FakeHandler.Http((_, _) => {
            var response = new HttpResponseMessage(HttpStatusCode.TemporaryRedirect);
            response.Headers.Location = new Uri("http://elsewhere.test/");
            return Task.FromResult(response);
        });
        using var client = new IppPrintClient(Printer, handler);

        IppException error = await Assert.ThrowsAsync<IppException>(() => client.GetCapabilitiesAsync());
        Assert.Contains("elsewhere.test", error.Message);

        // A caller-supplied handler that follows redirects itself is caught by the answering URI;
        // the document may have reached the other host, so a print is reported as uncertain.
        var following = FakeHandler.Http(async (request, _) => {
            await request.Content!.ReadAsByteArrayAsync();
            return new HttpResponseMessage(HttpStatusCode.OK) { RequestMessage = new HttpRequestMessage(HttpMethod.Post, "http://elsewhere.test/ipp/print") };
        });
        using var followingClient = new IppPrintClient(Printer, following);
        PdfPrintDeliveryException uncertain = await Assert.ThrowsAsync<PdfPrintDeliveryException>(() => followingClient.PrintAsync(new byte[8], "application/pdf"));
        Assert.Contains("elsewhere.test", uncertain.InnerException!.Message);
    }

    [Fact]
    public async Task RejectsOversizedResponsesAndMismatchedRequestIds() {
        var large = FakeHandler.Http((_, _) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(new byte[4096]) }));
        using (var client = new IppPrintClient(Printer, large, new IppPrintClientOptions { MaxResponseBytes = 2048 })) {
            await Assert.ThrowsAsync<IppException>(() => client.GetCapabilitiesAsync());
        }

        var mismatched = FakeHandler.Ipp(async request => {
            IppMessage response = Reply(await Decode(request), 0, null);
            response.RequestId++;
            return response;
        });
        using (var client = new IppPrintClient(Printer, mismatched)) {
            await Assert.ThrowsAsync<IppException>(() => client.CancelJobAsync(3));
        }
    }

    [Fact]
    public async Task TimesOutWithAProtocolException() {
        var handler = FakeHandler.Http(async (_, token) => {
            await Task.Delay(Timeout.Infinite, token);
            throw new InvalidOperationException();
        });
        using var client = new IppPrintClient(Printer, handler, new IppPrintClientOptions { Timeout = TimeSpan.FromMilliseconds(200) });

        IppException error = await Assert.ThrowsAsync<IppException>(() => client.GetJobAsync(1));
        Assert.IsType<TimeoutException>(error.InnerException);
    }

    [Fact]
    public async Task FallsBackToIpp11WhenThePrinterRejectsVersion2() {
        var versions = new List<string>();
        var handler = FakeHandler.Ipp(async request => {
            IppMessage message = await Decode(request);
            versions.Add($"{message.VersionMajor}.{message.VersionMinor}");
            return message.VersionMajor == 2 ? Reply(message, 0x0503, null) : Reply(message, 0, job => job.Add("job-id", IppValue.OfInteger(4)).Add("job-state", IppValue.OfEnum(9)));
        });
        using var client = new IppPrintClient(Printer, handler);

        IppJobInfo job = await client.GetJobAsync(4);
        await client.GetJobAsync(4);

        Assert.True(job.IsTerminal);
        Assert.Equal(new[] { "2.0", "1.1", "1.1" }, versions);
    }

    [Fact]
    public async Task MapsPrinterAttributesToCapabilities() {
        var handler = FakeHandler.Ipp(async request => Reply(await Decode(request), 0, null, printer => printer
            .Add("printer-make-and-model", IppValue.OfString(IppValueTag.TextWithoutLanguage, "Example"))
            .Add("printer-state", IppValue.OfEnum(5))
            .Add("printer-state-reasons", IppValue.OfString(IppValueTag.Keyword, "media-empty-error"))
            .Add("document-format-supported", IppValue.OfString(IppValueTag.MimeMediaType, "application/pdf"), IppValue.OfString(IppValueTag.MimeMediaType, "image/pwg-raster"))
            .Add("copies-supported", IppValue.OfRange(1, 99))
            .Add("printer-resolution-supported", IppValue.OfResolution(300, 300, 3), IppValue.OfResolution(600, 1200, 3))
            .Add("media-col-database",
                IppAttributeMapper.MediaCol(new IppMedia(21000, 29700) { SizeName = "iso_a4_210x297mm", LeftMargin = 0 }),
                IppValue.OfCollection(new[] { new IppAttribute("media-size", IppValue.OfCollection(new[] {
                    new IppAttribute("x-dimension", IppValue.OfRange(7620, 21590)),
                    new IppAttribute("y-dimension", IppValue.OfRange(12700, 35560)) })) }))
            .Add("orientation-requested-supported", IppValue.OfEnum(3), IppValue.OfEnum(4), IppValue.OfEnum(99))));
        using var client = new IppPrintClient(Printer, handler);

        IppPrinterCapabilities capabilities = await client.GetCapabilitiesAsync();

        Assert.Equal("Example", capabilities.MakeAndModel);
        Assert.Equal(IppPrinterState.Stopped, capabilities.State);
        Assert.Equal(new[] { "media-empty-error" }, capabilities.StateReasons);
        Assert.Equal(new[] { "application/pdf", "image/pwg-raster" }, capabilities.DocumentFormats);
        Assert.Equal(99, capabilities.MaxCopies);
        Assert.Equal(new IppResolution(600, 1200, IppResolutionUnits.DotsPerInch), capabilities.Resolutions[1]);
        IppMedia a4 = Assert.Single(capabilities.MediaSizes);
        Assert.Equal(("iso_a4_210x297mm", 21000, 29700, (int?)0), (a4.SizeName, a4.Width, a4.Height, a4.LeftMargin));
        Assert.Equal(new[] { IppOrientation.Portrait, IppOrientation.Landscape }, capabilities.Orientations);
    }

    [Fact]
    public async Task ValidatesUrisAndJobOptionsBeforeSending() {
        Assert.Throws<ArgumentException>(() => new IppPrintClient(new Uri("http://printer.test/ipp/print")));
        Assert.Throws<ArgumentException>(() => new IppPrintClient(new Uri("ipp://user:secret@printer.test/ipp/print")));
        using var client = new IppPrintClient(Printer, FakeHandler.Http((_, _) => throw new InvalidOperationException("must not send")));
        await Assert.ThrowsAsync<ArgumentException>(() => client.ValidateJobAsync("application/pdf", new IppJobOptions { Media = "iso_a4_210x297mm", MediaCol = new IppMedia(1, 1) }));
        await Assert.ThrowsAsync<ArgumentException>(() => client.ValidateJobAsync("application/pdf", new IppJobOptions { Sides = "two sided" }));
        await Assert.ThrowsAsync<ArgumentException>(() => client.ValidateJobAsync("application/pdf", new IppJobOptions { PageRanges = new[] { new IppPageRange(3, 4), new IppPageRange(1, 2) } }));
        await Assert.ThrowsAsync<ArgumentException>(() => client.ValidateJobAsync("pdf"));
    }

    private static async Task<IppMessage> Decode(HttpRequestMessage request) {
        byte[] body = await request.Content!.ReadAsByteArrayAsync();
        return IppCodec.Decode(body, body.Length, null, out _);
    }

    private static IppMessage Reply(IppMessage request, short status, Action<IppAttributeGroup>? job, Action<IppAttributeGroup>? printer = null) {
        var response = new IppMessage(request.VersionMajor, request.VersionMinor, status, request.RequestId);
        response.AddGroup(IppGroupTag.Operation)
            .Add("attributes-charset", IppValue.OfString(IppValueTag.Charset, "utf-8"))
            .Add("attributes-natural-language", IppValue.OfString(IppValueTag.NaturalLanguage, "en"));
        if (job is not null) job(response.AddGroup(IppGroupTag.Job));
        if (printer is not null) printer(response.AddGroup(IppGroupTag.Printer));
        return response;
    }

    private sealed class FakeHandler : HttpMessageHandler {
        private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _send;

        private FakeHandler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) => _send = send;

        public static FakeHandler Http(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> send) => new(send);

        public static FakeHandler Ipp(Func<HttpRequestMessage, Task<IppMessage>> send) => new(async (request, _) => {
            var content = new ByteArrayContent(IppCodec.Encode(await send(request)));
            content.Headers.ContentType = new MediaTypeHeaderValue("application/ipp");
            return new HttpResponseMessage(HttpStatusCode.OK) { Content = content };
        });

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => _send(request, cancellationToken);
    }

    private sealed class NonSeekableStream : MemoryStream {
        public NonSeekableStream(byte[] data) : base(data) { }
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => base.Position; set => throw new NotSupportedException(); }
    }
}
