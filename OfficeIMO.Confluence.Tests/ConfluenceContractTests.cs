using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Confluence;
using Xunit;

namespace OfficeIMO.Confluence.Tests;

public sealed class ConfluenceContractTests {
    [Fact]
    public void Session_RejectsCredentialsEmbeddedInSiteUri() {
        var options = new ConfluenceSessionOptions { SiteUri = new Uri("https://user:secret@example.atlassian.net/") };

        Assert.Throws<ArgumentException>(() => new ConfluenceSession(new ConfluenceBearerCredentialSource("token"), options));
    }

    [Fact]
    public async Task PageBatch_ExposesDecodedNextCursor() {
        var handler = new RecordingHandler(_ => Response(HttpStatusCode.OK, "{\"results\":[],\"_links\":{\"next\":\"/wiki/api/v2/pages?limit=25&cursor=abc%2B123\"}}"));
        using ConfluenceClient client = CreateClient(handler);

        ConfluencePageBatch batch = await client.GetPagesAsync();

        Assert.Equal("abc+123", batch.NextCursor);
    }

    [Fact]
    public async Task GetPage_UsesV2BodyFormatAndCallerCredentialSource() {
        string fixture = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "page-adf.json"));
        var handler = new RecordingHandler(_ => Response(HttpStatusCode.OK, fixture));
        using ConfluenceClient client = CreateClient(handler, new ConfluenceBasicCredentialSource("person@example.com", "token"));

        ConfluencePage page = await client.GetPageAsync("123", ConfluenceBodyFormat.AtlasDocFormat);

        Assert.Equal("Status", page.Title);
        Assert.Equal("/wiki/api/v2/pages/123?body-format=atlas_doc_format", handler.Requests.Single().RequestUri!.PathAndQuery);
        Assert.Equal("Basic", handler.Requests.Single().Headers.Authorization!.Scheme);
    }

    [Fact]
    public void UpdatePlan_IsDryAndCarriesOptimisticVersion() {
        ConfluencePageBody body = ConfluenceContentConverter.FromMarkdown("# Ready").Value;

        ConfluencePageWritePlan plan = ConfluenceClient.PlanUpdatePage(new ConfluencePageUpdateRequest {
            PageId = "123",
            Title = "Ready",
            VersionNumber = 8,
            VersionMessage = "generated",
            Body = body,
        });

        Assert.Equal("PUT", plan.Method);
        Assert.Equal("/wiki/api/v2/pages/123", plan.RelativeUri);
        Assert.Contains("\"number\":8", plan.Payload);
        Assert.Contains("atlas_doc_format", plan.Payload);
    }

    [Fact]
    public async Task NonIdempotentWrite_IsNotRetriedAfterServiceFailure() {
        var handler = new RecordingHandler(_ => Response(HttpStatusCode.ServiceUnavailable, "{\"message\":\"later\"}"));
        using ConfluenceClient client = CreateClient(handler);
        var request = new ConfluencePageCreateRequest {
            SpaceId = "42",
            Title = "Status",
            Body = ConfluenceContentConverter.FromMarkdown("Ready").Value,
        };

        await Assert.ThrowsAsync<ConfluenceApiException>(() => client.CreatePageAsync(request));
        Assert.Single(handler.Requests);
    }

    [Fact]
    public async Task SafeRead_RetriesRateLimit() {
        int attempt = 0;
        string fixture = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "page-adf.json"));
        var handler = new RecordingHandler(_ => ++attempt == 1 ? Response((HttpStatusCode)429, "{}") : Response(HttpStatusCode.OK, fixture));
        using ConfluenceClient client = CreateClient(handler, configure: options => { options.RetryBaseDelay = TimeSpan.Zero; options.RetryMaxDelay = TimeSpan.Zero; });

        ConfluencePage page = await client.GetPageAsync("123");

        Assert.Equal("123", page.Id);
        Assert.Equal(2, handler.Requests.Count);
    }

    [Fact]
    public void ManagedSection_ReplacesOnlyMarkedContentAndProducesHashes() {
        string original = "<p>Owner content</p>\n<!-- officeimo:section:report:start -->\nold\n<!-- officeimo:section:report:end -->\n<p>Tail</p>";

        ConfluenceManagedSectionResult result = ConfluenceManagedSection.Apply(original, "report", "<p>new</p>");

        Assert.Contains("<p>Owner content</p>", result.UpdatedBody);
        Assert.Contains("<p>new</p>", result.UpdatedBody);
        Assert.Contains("<p>Tail</p>", result.UpdatedBody);
        Assert.True(result.Changed);
        Assert.NotEqual(result.OriginalSha256, result.UpdatedSha256);
    }

    [Fact]
    public async Task AttachmentUpload_UsesMultipartAndXsrfHeader() {
        var handler = new RecordingHandler(request => {
            Assert.Equal(HttpMethod.Put, request.Method);
            Assert.True(request.Headers.Contains("X-Atlassian-Token"));
            Assert.NotNull(request.Content);
            Assert.StartsWith("multipart/form-data", request.Content!.Headers.ContentType!.MediaType);
            return Response(HttpStatusCode.OK, "{\"results\":[{\"id\":\"a1\",\"title\":\"report.txt\",\"extensions\":{\"mediaType\":\"text/plain\",\"fileSize\":5},\"_links\":{\"download\":\"/download/attachments/123/report.txt\"}}]}" );
        });
        using ConfluenceClient client = CreateClient(handler);

        IReadOnlyList<ConfluenceAttachment> result = await client.UploadAttachmentAsync("123", new ConfluenceAttachmentUpload {
            FileName = "report.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("ready"),
        });

        ConfluenceAttachment attachment = Assert.Single(result);
        Assert.Equal("a1", attachment.Id);
        Assert.Equal("123", attachment.PageId);
        Assert.Equal("text/plain", attachment.MediaType);
        Assert.Equal(5, attachment.FileSize);
        Assert.Equal("/download/attachments/123/report.txt", attachment.DownloadLink);
    }

    private static ConfluenceClient CreateClient(RecordingHandler handler, IConfluenceCredentialSource? credential = null, Action<ConfluenceSessionOptions>? configure = null) {
        var options = new ConfluenceSessionOptions {
            SiteUri = new Uri("https://example.atlassian.net/"),
            HttpClient = new HttpClient(handler),
        };
        configure?.Invoke(options);
        return new ConfluenceSession(credential ?? new ConfluenceBearerCredentialSource("token"), options).CreateClient();
    }

    private static HttpResponseMessage Response(HttpStatusCode status, string body) => new HttpResponseMessage(status) {
        Content = new StringContent(body, Encoding.UTF8, "application/json"),
    };

    private sealed class RecordingHandler : HttpMessageHandler {
        private readonly Func<HttpRequestMessage, HttpResponseMessage> _respond;
        internal RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> respond) => _respond = respond;
        internal List<HttpRequestMessage> Requests { get; } = new List<HttpRequestMessage>();

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
            var snapshot = new HttpRequestMessage(request.Method, request.RequestUri);
            snapshot.Headers.Authorization = request.Headers.Authorization == null ? null : new AuthenticationHeaderValue(request.Headers.Authorization.Scheme, request.Headers.Authorization.Parameter);
            foreach (var header in request.Headers.Where(header => !string.Equals(header.Key, "Authorization", StringComparison.OrdinalIgnoreCase))) snapshot.Headers.TryAddWithoutValidation(header.Key, header.Value);
            snapshot.Content = request.Content;
            Requests.Add(snapshot);
            return Task.FromResult(_respond(request));
        }
    }
}
