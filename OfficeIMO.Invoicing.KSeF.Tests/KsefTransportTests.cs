using System.Net;
using OfficeIMO.Invoicing.KSeF;

namespace OfficeIMO.Invoicing.KSeF.Tests;

public class KsefTransportTests {
    [Theory]
    [InlineData("redirect")]
    [InlineData("too-large")]
    [InlineData("malformed")]
    [InlineData("quota")]
    public async Task ReadOnlyTransportBoundsAndRejectionsDoNotTriggerRetries(string scenario) {
        int calls = 0;
        using var handler = new ResponseHandler(request => {
            calls++;
            var response = new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(scenario == "malformed" ? "{" : "[]") };
            if (scenario == "redirect") { response.StatusCode = HttpStatusCode.TemporaryRedirect; response.Headers.Location = new Uri("https://untrusted.invalid/"); }
            if (scenario == "too-large") response.Content.Headers.ContentLength = 1024 * 1024 + 1;
            if (scenario == "quota") { response.StatusCode = HttpStatusCode.TooManyRequests; response.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(15)); }
            return response;
        });
        using var client = new KsefClient(handler: handler, clock: new KsefClock());
        if (scenario is "redirect" or "quota") {
            KsefApiException error = await Assert.ThrowsAsync<KsefApiException>(() => client.CheckEncryptionKeyAsync());
            if (scenario == "quota") Assert.Equal(TimeSpan.FromSeconds(15), error.RetryAfter);
        } else await Assert.ThrowsAsync<InvalidDataException>(() => client.CheckEncryptionKeyAsync());
        Assert.Equal(1, calls);
    }
    private sealed class ResponseHandler(Func<HttpRequestMessage, HttpResponseMessage> response) : HttpMessageHandler {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => Task.FromResult(response(request));
    }
}
