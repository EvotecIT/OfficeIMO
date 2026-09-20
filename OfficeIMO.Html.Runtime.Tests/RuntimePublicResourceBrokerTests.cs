using System.Net;
using System.Net.Sockets;
using System.Text;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Tests;
using Xunit;

namespace OfficeIMO.Html.Runtime.Tests;

public class RuntimePublicResourceBrokerTests {
    [Theory]
    [InlineData("0.1.2.3")]
    [InlineData("10.1.2.3")]
    [InlineData("100.64.1.2")]
    [InlineData("127.0.0.1")]
    [InlineData("169.254.1.2")]
    [InlineData("172.16.0.1")]
    [InlineData("192.0.0.1")]
    [InlineData("192.0.2.10")]
    [InlineData("192.88.99.1")]
    [InlineData("192.168.1.1")]
    [InlineData("198.18.0.1")]
    [InlineData("198.51.100.1")]
    [InlineData("203.0.113.1")]
    [InlineData("224.0.0.1")]
    [InlineData("255.255.255.255")]
    [InlineData("::1")]
    public void NonPublicAndUnsupportedAddressesAreRejected(string address) =>
        Assert.False(HtmlPublicResourceBroker.IsPublicIpv4(IPAddress.Parse(address)));

    [Theory]
    [InlineData("1.1.1.1")]
    [InlineData("8.8.8.8")]
    [InlineData("93.184.215.14")]
    public void PublicIpv4AddressesAreAdmitted(string address) =>
        Assert.True(HtmlPublicResourceBroker.IsPublicIpv4(IPAddress.Parse(address)));

    [Fact]
    public void MixedDnsAnswerCannotChoosePublicAddressAroundPrivateOne() {
        var addresses = new[] { IPAddress.Parse("1.1.1.1"), IPAddress.Parse("127.0.0.1") };
        Assert.Throws<HtmlScriptRuntimeException>(() => HtmlPublicResourceBroker.SelectPublicAddress(addresses));
        Assert.Equal(IPAddress.Parse("1.1.1.1"), HtmlPublicResourceBroker.SelectPublicAddress(
            new[] { IPAddress.Parse("2606:4700:4700::1111"), IPAddress.Parse("1.1.1.1") }));
    }

    [Fact]
    public void RedirectsRequireAnAllowedHostAndCannotDowngradeTls() {
        var broker = new HtmlPublicResourceBroker(new[] { "example.com" });
        var initial = new Uri("https://example.com/report");
        Assert.Equal(new Uri("https://example.com/final"), broker.ValidateRedirect(initial, new Uri("https://example.com/final")));
        Assert.Throws<HtmlScriptRuntimeException>(() => broker.ValidateRedirect(initial, new Uri("http://example.com/final")));
        Assert.Throws<HtmlScriptRuntimeException>(() => broker.ValidateRedirect(initial, new Uri("https://other.example/final")));
        Assert.Throws<ArgumentException>(() => broker.ValidateRedirect(initial, new Uri("https://example.com:8443/final")));
        Assert.Throws<ArgumentException>(() => HtmlPublicResourceBroker.ValidateUrl(new Uri("http://user:pass@example.com/")));
        Assert.Equal(new Uri("https://example.com/final#review"), HtmlPublicResourceBroker.ResolveRedirect(
            new Uri("https://example.com/report#review"), new Uri("/final", UriKind.Relative)));
        Assert.Equal(new Uri("https://example.com/final#other"), HtmlPublicResourceBroker.ResolveRedirect(
            new Uri("https://example.com/report#review"), new Uri("/final#other", UriKind.Relative)));
    }

    [Fact]
    public void ApprovedHostsExposeOnlyStandardHttpOriginsToTheOfflineWorker() {
        var broker = new HtmlPublicResourceBroker(new[] { "example.com", "assets.example.com" });

        Assert.Equal(new[] {
            "http://example.com/", "https://example.com/",
            "http://assets.example.com/", "https://assets.example.com/"
        }.Order(), broker.AllowedOrigins.Select(origin => origin.AbsoluteUri).Order());
        Assert.True(broker.AllowsHost(new Uri("https://assets.example.com/theme.css")));
        Assert.False(broker.AllowsHost(new Uri("https://other.example.com/theme.css")));
    }

    [Fact]
    public void HtmlDecodingRejectsInvalidUtf8AndUnqualifiedMediaTypes() {
        var url = new Uri("https://example.com/");
        Assert.Equal("<p>Ready</p>", HtmlPublicResourceBroker.DecodeUtf8Html(
            HtmlRuntimeResource.FromText(url, "<p>Ready</p>", "text/html; charset=utf-8")));
        Assert.Throws<HtmlScriptRuntimeException>(() => HtmlPublicResourceBroker.DecodeUtf8Html(
            new HtmlRuntimeResource(url, new byte[] { 0xC3, 0x28 }, "text/html")));
        Assert.Throws<HtmlScriptRuntimeException>(() => HtmlPublicResourceBroker.DecodeUtf8Html(
            HtmlRuntimeResource.FromText(url, "<p>Ready</p>", "text/html; charset=windows-1252")));
        Assert.Throws<HtmlScriptRuntimeException>(() => HtmlPublicResourceBroker.DecodeUtf8Html(
            HtmlRuntimeResource.FromText(url, "<p>Ready</p>", "application/octet-stream")));
    }

    [Fact]
    public async Task LoopbackLiteralIsDeniedBeforeHttpConnection() {
        var broker = new HtmlPublicResourceBroker(new[] { "127.0.0.1" });
        Exception error = await Assert.ThrowsAnyAsync<Exception>(() => broker.FetchAsync(new Uri("http://127.0.0.1/")));
        Assert.Contains("non-public address", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task AcquisitionRevalidatesDnsAndStopsSameHostRebindingAfterRedirect() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/start"
            ? RuntimeHttpFixture.Reply.Text("", "text/plain", 302, "Location: /final\r\n")
            : RuntimeHttpFixture.Reply.Text("unexpected", "text/plain")));
        int resolutions = 0, connections = 0;
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) => Task.FromResult(
            new[] { IPAddress.Parse(Interlocked.Increment(ref resolutions) == 1 ? "93.184.216.34" : "127.0.0.1") }),
            (address, port) => {
                Assert.Equal(IPAddress.Parse("93.184.216.34"), address);
                Assert.Equal(80, port);
                Interlocked.Increment(ref connections);
            });

        Exception error = await Assert.ThrowsAnyAsync<Exception>(() =>
            broker.FetchAsync(new Uri("http://page.example.test/start")));

        Assert.Contains("non-public address", error.ToString(), StringComparison.Ordinal);
        Assert.Equal(2, resolutions);
        Assert.Equal(1, connections);
        Assert.Equal(new[] { "/start" }, server.Requests);
    }

    [Fact]
    public async Task AcquisitionRecordsSameHostAndApprovedCrossHostRedirects() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path switch {
            "/same" => RuntimeHttpFixture.Reply.Text("", "text/plain", 302, "Location: /final\r\n"),
            "/cross" => RuntimeHttpFixture.Reply.Text("", "text/plain", 302,
                "Location: http://assets.example.test/final\r\n"),
            _ => RuntimeHttpFixture.Reply.Text("ready", "text/plain")
        }));
        int resolutions = 0;
        var connections = new List<(IPAddress Address, int Port)>();
        var broker = Broker(server, new[] { "page.example.test", "assets.example.test" }, (host, _) => {
            Interlocked.Increment(ref resolutions);
            return Task.FromResult(new[] { IPAddress.Parse(host.StartsWith("assets", StringComparison.Ordinal) ?
                "1.1.1.1" : "93.184.216.34") });
        }, (address, port) => connections.Add((address, port)));

        HtmlPublicResourceResult same = await broker.FetchAsync(new Uri("http://page.example.test/same"));
        HtmlPublicResourceResult cross = await broker.FetchAsync(new Uri("http://page.example.test/cross"));

        Assert.Equal(new Uri("http://page.example.test/final"), same.Resource.FinalUrl);
        Assert.Equal(new Uri("http://assets.example.test/final"), cross.Resource.FinalUrl);
        Assert.Equal(302, Assert.Single(same.Redirects).StatusCode);
        HtmlPublicRedirect hop = Assert.Single(cross.Redirects);
        Assert.Equal(IPAddress.Parse("93.184.216.34"), hop.ConnectedAddress);
        Assert.Equal("ready", Encoding.UTF8.GetString(cross.Resource.Content));
        Assert.Equal(IPAddress.Parse("1.1.1.1"), cross.ConnectedAddress);
        Assert.Equal(4, resolutions);
        Assert.Equal(new[] {
            (IPAddress.Parse("93.184.216.34"), 80),
            (IPAddress.Parse("93.184.216.34"), 80),
            (IPAddress.Parse("93.184.216.34"), 80),
            (IPAddress.Parse("1.1.1.1"), 80)
        }, connections);
    }

    [Fact]
    public async Task DynamicAcquisitionSendsExactRequestAndRetainsNonSuccessResponse() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("conflict", "text/plain", 409, "X-Result: retained\r\n")));
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }));
        var request = new HtmlRuntimeFetchRequest(new Uri("http://page.example.test/submit"),
            new Uri("http://page.example.test/"), "POST",
            new Dictionary<string, string> { ["Content-Type"] = "application/json", ["X-Variant"] = "blue" },
            Encoding.UTF8.GetBytes("{\"value\":42}"), credentials: "omit");

        HtmlPublicResourceResult result = await broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1));

        Assert.Equal(409, result.Resource.StatusCode);
        Assert.Equal("retained", result.Resource.Headers["X-Result"]);
        Assert.Equal("conflict", Encoding.UTF8.GetString(result.Resource.Content));
        Assert.Equal(1, result.DynamicRequest!.Occurrence);
        RuntimeHttpFixture.ReceivedRequest received = Assert.Single(server.Received);
        Assert.Equal("POST", received.Method);
        Assert.Equal("blue", received.Headers["X-Variant"]);
        Assert.Equal("{\"value\":42}", Encoding.UTF8.GetString(received.Body));
    }

    [Fact]
    public async Task DynamicAcquisitionRejectsCredentialsAndHonorsRedirectLimit() {
        var credentialed = new HtmlRuntimeFetchRequest(new Uri("http://page.example.test/private"),
            new Uri("http://page.example.test/"), headers:
            new Dictionary<string, string> { ["Authorization"] = "Bearer secret" });
        Assert.Throws<HtmlScriptRuntimeException>(() => HtmlPublicResourceBroker.ValidateDynamicRequest(credentialed));

        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("", "text/plain", 302, "Location: /final\r\n")));
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }));
        var request = new HtmlRuntimeFetchRequest(new Uri("http://page.example.test/submit"),
            new Uri("http://page.example.test/"), "POST", body: Encoding.UTF8.GetBytes("once"));

        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1)));
        Assert.Contains("redirect limit", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(6, server.Requests.Count);
    }

    [Fact]
    public async Task DynamicPostRedirectChangesToGetAndRetainsDirectHops() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/submit"
            ? RuntimeHttpFixture.Reply.Text("", "text/plain", 302, "Location: /result\r\n")
            : RuntimeHttpFixture.Reply.Text("ready", "text/plain")));
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }));
        var request = new HtmlRuntimeFetchRequest(new Uri("http://page.example.test/submit"),
            new Uri("http://page.example.test/"), "POST", body: Encoding.UTF8.GetBytes("once"));

        HtmlPublicResourceResult result = await broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1));

        Assert.Equal("ready", Encoding.UTF8.GetString(result.Resource.Buffer));
        Assert.Equal(2, result.DynamicHops!.Count);
        Assert.Equal(302, result.DynamicHops[0].Response.StatusCode);
        Assert.Equal(new[] { "POST", "GET" }, server.Received.Select(item => item.Method));
        Assert.Empty(server.Received.Last().Body);
    }

    [Fact]
    public async Task Dynamic307RetainsPostBodyAndChargesEveryHop() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/submit"
            ? RuntimeHttpFixture.Reply.Text("", "text/plain", 307, "Location: /result\r\n")
            : RuntimeHttpFixture.Reply.Text("ready", "text/plain")));
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (_, _, token) => {
                var client = new TcpClient(AddressFamily.InterNetwork);
                try { await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token); return client.GetStream(); }
                catch { client.Dispose(); throw; }
            }, maxRequestBytes: 4, maxTotalRequestBytes: 8);
        var request = new HtmlRuntimeFetchRequest(new Uri("http://page.example.test/submit"),
            new Uri("http://page.example.test/"), "POST", body: Encoding.UTF8.GetBytes("once"));

        HtmlPublicResourceResult result = await broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1));

        Assert.Equal(new[] { "POST", "POST" }, server.Received.Select(item => item.Method));
        Assert.All(server.Received, item => Assert.Equal("once", Encoding.UTF8.GetString(item.Body)));
        Assert.Equal(2, result.HttpExchanges!.Count);
        Assert.All(result.HttpExchanges, exchange => Assert.Equal(4, exchange.RequestBodyByteCount));
    }

    [Fact]
    public async Task DynamicRedirectToAnotherOriginStopsBeforeSecondConnection() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("", "text/plain", 302,
                "Location: http://other.example.test/result\r\n")));
        int connections = 0;
        var broker = Broker(server, new[] { "page.example.test", "other.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            (_, _) => Interlocked.Increment(ref connections));
        var request = new HtmlRuntimeFetchRequest(new Uri("http://page.example.test/start"),
            new Uri("http://page.example.test/"));

        HtmlScriptRuntimeException error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1)));

        Assert.Contains("distinct origins", error.Message, StringComparison.Ordinal);
        Assert.Equal(1, connections);
    }

    [Fact]
    public async Task NavigationRedirectUsesSeparateOriginAuthorityAndPostRewriteRules() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/submit"
            ? RuntimeHttpFixture.Reply.Text("", "text/html", 302,
                "Location: http://reports.example.test/result\r\n")
            : RuntimeHttpFixture.Reply.Text("<h1>ready</h1>", "text/html; charset=utf-8")));
        var navigationOrigins = new[] {
            new Uri("http://page.example.test/"), new Uri("http://reports.example.test/")
        };
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (_, _, token) => {
                var client = new TcpClient(AddressFamily.InterNetwork);
                try { await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token); return client.GetStream(); }
                catch { client.Dispose(); throw; }
            }, navigationOrigins: navigationOrigins);
        var request = new HtmlRuntimeNavigationRequest(new Uri("http://page.example.test/submit"),
            new Uri("http://page.example.test/start?private=1#fragment"),
            new Uri("http://page.example.test/start?private=1"), "POST",
            new Dictionary<string, string> { ["Content-Type"] = "text/plain" }, Encoding.UTF8.GetBytes("once"));

        HtmlPublicResourceResult result = await broker.FetchAsync(new HtmlRuntimeNavigationDiscovery(request, 1));

        RuntimeHttpFixture.ReceivedRequest[] received = server.Received.ToArray();
        Assert.Equal(new[] { "POST", "GET" }, received.Select(item => item.Method));
        Assert.Equal("once", Encoding.UTF8.GetString(received[0].Body));
        Assert.Empty(received[1].Body);
        Assert.Equal("http://page.example.test/start?private=1", received[0].Headers["Referer"]);
        Assert.Equal("http://page.example.test/", received[1].Headers["Referer"]);
        Assert.Equal(new Uri("http://reports.example.test/result"), result.Resource.FinalUrl);
        Assert.Equal(2, result.NavigationHops!.Count);
        Assert.Equal(new[] { "POST", "GET" }, result.HttpExchanges!.Select(exchange => exchange.Method));
    }

    [Fact]
    public async Task NavigationRedirectStopsBeforeAnUnapprovedOriginConnection() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("", "text/html", 302,
                "Location: http://other.example.test/result\r\n")));
        int connections = 0;
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (_, _, token) => {
                Interlocked.Increment(ref connections);
                var client = new TcpClient(AddressFamily.InterNetwork);
                try { await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token); return client.GetStream(); }
                catch { client.Dispose(); throw; }
            }, navigationOrigins: new[] { new Uri("http://page.example.test/") });
        var request = new HtmlRuntimeNavigationRequest(new Uri("http://page.example.test/start"),
            new Uri("http://page.example.test/"));

        HtmlScriptRuntimeException error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new HtmlRuntimeNavigationDiscovery(request, 1)));

        Assert.Contains("redirect origin was not authorized", error.Message, StringComparison.Ordinal);
        Assert.Equal(1, connections);
    }

    [Fact]
    public async Task NavigationRedirectDoesNotRestoreASuppressedReferrer() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/start"
            ? RuntimeHttpFixture.Reply.Text("", "text/html", 302,
                "Location: /result\r\nReferrer-Policy: no-referrer\r\n")
            : RuntimeHttpFixture.Reply.Text("<h1>ready</h1>", "text/html")));
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (_, _, token) => {
                var client = new TcpClient(AddressFamily.InterNetwork);
                try { await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token); return client.GetStream(); }
                catch { client.Dispose(); throw; }
            }, navigationOrigins: new[] { new Uri("http://page.example.test/") });
        var request = new HtmlRuntimeNavigationRequest(new Uri("http://page.example.test/start"),
            new Uri("http://page.example.test/private?token=hidden"),
            new Uri("http://page.example.test/private?token=hidden"));

        await broker.FetchAsync(new HtmlRuntimeNavigationDiscovery(request, 1));

        RuntimeHttpFixture.ReceivedRequest[] received = server.Received.ToArray();
        Assert.Equal("http://page.example.test/private?token=hidden", received[0].Headers["Referer"]);
        Assert.False(received[1].Headers.ContainsKey("Referer"));
        Assert.Null(HtmlPublicResourceBroker.ApplyRedirectReferrerPolicy(null,
            new Uri("http://page.example.test/result"), new Dictionary<string, string>()));
    }

    [Fact]
    public async Task DynamicAcquisitionUsesRetainedInitiatorBeforeSendingASideEffect() {
        int connections = 0;
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            (_, _, _) => { Interlocked.Increment(ref connections); return ValueTask.FromResult<Stream>(Stream.Null); });
        var request = new HtmlRuntimeFetchRequest(new Uri("http://page.example.test/submit"),
            new Uri("http://other.example.test/"), "POST", body: Encoding.UTF8.GetBytes("once"));

        HtmlScriptRuntimeException error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1)));

        Assert.Contains("origin was not authorized", error.Message, StringComparison.Ordinal);
        Assert.Equal(0, connections);
    }

    [Fact]
    public async Task CrossOriginDynamicAcquisitionPreflightsBeforeUnsafeRequest() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ready", "text/plain")));
        server.RespondToRequest = (received, _) => Task.FromResult(received.Method == "OPTIONS"
            ? RuntimeHttpFixture.Reply.Text("", "text/plain", 204,
                "Access-Control-Allow-Origin: http://page.example.test\r\nAccess-Control-Allow-Methods: POST\r\nAccess-Control-Allow-Headers: content-type\r\n")
            : RuntimeHttpFixture.Reply.Text("ready", "text/plain", 200,
                "Access-Control-Allow-Origin: http://page.example.test\r\n"));
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            dynamicOrigins: new[] { new Uri("http://api.example.test/") });
        var request = new HtmlRuntimeFetchRequest(new Uri("http://api.example.test/submit"),
            new Uri("http://page.example.test/"), "POST",
            new Dictionary<string, string> { ["Content-Type"] = "application/json" }, Encoding.UTF8.GetBytes("{}"),
            credentials: "same-origin");

        HtmlPublicResourceResult result = await broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1));

        Assert.Equal(new[] { "OPTIONS", "POST" }, server.Received.Select(item => item.Method));
        Assert.NotNull(Assert.Single(result.DynamicHops!).PreflightResponse);
        Assert.Equal("http://page.example.test", server.Received.First().Headers["Origin"]);
        Assert.Equal("POST", server.Received.First().Headers["Access-Control-Request-Method"]);
        Assert.Equal("content-type", server.Received.First().Headers["Access-Control-Request-Headers"]);
        Assert.Equal("http://page.example.test", server.Received.Last().Headers["Origin"]);
    }

    [Fact]
    public async Task DynamicOnlyHostCannotBeFetchedAsAStaticResource() {
        await using var server = new RuntimeHttpFixture((path, _) => Task.FromResult(path == "/start"
            ? RuntimeHttpFixture.Reply.Text("", "text/plain", 302,
                "Location: http://api.example.test/data\r\n")
            : RuntimeHttpFixture.Reply.Text("ready", "text/plain", 200,
                "Access-Control-Allow-Origin: http://page.example.test\r\n")));
        int connections = 0;
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (_, _, token) => {
                Interlocked.Increment(ref connections);
                var client = new TcpClient(AddressFamily.InterNetwork);
                try { await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token); return client.GetStream(); }
                catch { client.Dispose(); throw; }
            }, dynamicOrigins: new[] { new Uri("http://api.example.test/") });

        Assert.False(broker.AllowsHost(new Uri("http://api.example.test/data")));
        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new Uri("http://api.example.test/data")));
        Assert.Equal(0, connections);

        HtmlScriptRuntimeException redirectError = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new Uri("http://page.example.test/start")));
        Assert.Contains("redirect host is not allowed", redirectError.Message, StringComparison.Ordinal);
        Assert.Equal(1, connections);

        var dynamicRequest = new HtmlRuntimeFetchRequest(new Uri("http://api.example.test/data"),
            new Uri("http://page.example.test/"));
        HtmlPublicResourceResult dynamicResult = await broker.FetchAsync(
            new HtmlRuntimeFetchDiscovery(dynamicRequest, 1));
        Assert.Equal("ready", Encoding.UTF8.GetString(dynamicResult.Resource.Content));
        Assert.Equal(2, connections);
        Assert.Equal(new[] { "/start", "/data" }, server.Requests);
    }

    [Fact]
    public async Task FailedPreflightDoesNotSendUnsafeRequest() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("", "text/plain", 204)));
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            dynamicOrigins: new[] { new Uri("http://api.example.test/") });
        var request = new HtmlRuntimeFetchRequest(new Uri("http://api.example.test/submit"),
            new Uri("http://page.example.test/"), "POST",
            new Dictionary<string, string> { ["Content-Type"] = "application/json" }, Encoding.UTF8.GetBytes("{}"),
            credentials: "same-origin");

        await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => broker.FetchAsync(
            new HtmlRuntimeFetchDiscovery(request, 1)));

        Assert.Equal("OPTIONS", Assert.Single(server.Received).Method);
    }

    [Fact]
    public async Task DynamicAcquisitionCancellationStopsWithoutRetryingThePost() {
        var arrived = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("late", "text/plain")));
        server.RespondToRequest = async (_, token) => {
            arrived.TrySetResult();
            await Task.Delay(TimeSpan.FromSeconds(5), token);
            return RuntimeHttpFixture.Reply.Text("late", "text/plain");
        };
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }));
        var request = new HtmlRuntimeFetchRequest(new Uri("http://page.example.test/submit"),
            new Uri("http://page.example.test/"), "POST", body: Encoding.UTF8.GetBytes("once"));
        using var cancellation = new CancellationTokenSource();
        Task<HtmlPublicResourceResult> acquisition = broker.FetchAsync(new HtmlRuntimeFetchDiscovery(request, 1),
            cancellation.Token);
        await arrived.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => acquisition);

        Assert.Single(server.Received);
    }

    [Fact]
    public async Task AcquisitionRejectsUnapprovedCrossHostRedirectBeforeSecondConnection() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("", "text/plain", 302,
                "Location: http://other.example.test/final\r\n")));
        int connections = 0;
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            (_, _) => Interlocked.Increment(ref connections));

        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new Uri("http://page.example.test/start")));

        Assert.Contains("redirect host is not allowed", error.Message, StringComparison.Ordinal);
        Assert.Equal(1, connections);
        Assert.Equal(new[] { "/start" }, server.Requests);
    }

    [Fact]
    public async Task AcquisitionRejectsOversizedDeclaredResponseBeforeReadingIt() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(new RuntimeHttpFixture.Reply(
            new byte[4 * 1024 * 1024 + 1], "application/octet-stream")));
        var broker = Broker(server, new[] { "page.example.test" }, (_, _) =>
            Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }));

        var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new Uri("http://page.example.test/large")));

        Assert.Contains("byte budget", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AcquisitionHonorsAStricterPerRunResponseLimit() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("too large", "text/plain")));
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (_, _, token) => {
                var client = new TcpClient(AddressFamily.InterNetwork);
                try {
                    await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token);
                    return client.GetStream();
                } catch { client.Dispose(); throw; }
            }, maxRequests: 1, maxResourceBytes: 4, maxTotalBytes: 4);

        HtmlScriptRuntimeException error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new Uri("http://page.example.test/")));

        Assert.Contains("byte budget", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AcquisitionHonorsAStricterPerRunRedirectLimit() {
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(
            RuntimeHttpFixture.Reply.Text("", "text/plain", 302, "Location: /final\r\n")));
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (_, _, token) => {
                var client = new TcpClient(AddressFamily.InterNetwork);
                try {
                    await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token);
                    return client.GetStream();
                } catch { client.Dispose(); throw; }
            }, maxRedirects: 0);

        HtmlScriptRuntimeException error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() =>
            broker.FetchAsync(new Uri("http://page.example.test/start")));

        Assert.Contains("redirect limit", error.Message, StringComparison.Ordinal);
        Assert.Single(server.Requests);
    }

    [Fact]
    public async Task AcquisitionHonorsAStricterPerRunTimeout() {
        await using var server = new RuntimeHttpFixture(async (_, _) => {
            await Task.Delay(TimeSpan.FromSeconds(1));
            return RuntimeHttpFixture.Reply.Text("late", "text/plain");
        });
        var broker = new HtmlPublicResourceBroker(new[] { "page.example.test" },
            (_, _) => Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") }),
            async (_, _, token) => {
                var client = new TcpClient(AddressFamily.InterNetwork);
                try {
                    await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token);
                    return client.GetStream();
                } catch { client.Dispose(); throw; }
            }, timeout: TimeSpan.FromMilliseconds(100));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            broker.FetchAsync(new Uri("http://page.example.test/slow")));
    }

    [Fact]
    public async Task DefaultConnectorUsesExactEndpointAndOwnsSocket() {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        Stream? stream = null;
        Socket? accepted = null;
        try {
            stream = await HtmlPublicResourceBroker.ConnectToAddressAsync(
                IPAddress.Loopback, port, CancellationToken.None);
            accepted = await listener.AcceptSocketAsync();
            Assert.Equal(IPAddress.Loopback, ((IPEndPoint)accepted.RemoteEndPoint!).Address);

            await stream.DisposeAsync();
            stream = null;
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            int received = await accepted.ReceiveAsync(new Memory<byte>(new byte[1]), SocketFlags.None, timeout.Token);
            Assert.Equal(0, received);
        } finally {
            if (stream is not null) await stream.DisposeAsync();
            accepted?.Dispose();
            listener.Stop();
        }
    }

    [Fact]
    public async Task DefaultConnectorHonorsPreCancelledToken() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await HtmlPublicResourceBroker.ConnectToAddressAsync(IPAddress.Loopback, 9, cancellation.Token));
    }

    private static HtmlPublicResourceBroker Broker(RuntimeHttpFixture server, IEnumerable<string> hosts,
        Func<string, CancellationToken, Task<IPAddress[]>> resolve,
        Action<IPAddress, int>? connected = null, IEnumerable<Uri>? dynamicOrigins = null) =>
        new(hosts, resolve, async (address, port, token) => {
            connected?.Invoke(address, port);
            var client = new TcpClient(AddressFamily.InterNetwork);
            try {
                await client.ConnectAsync(IPAddress.Loopback, server.Origin.Port, token);
                return client.GetStream();
            } catch { client.Dispose(); throw; }
        }, dynamicOrigins: dynamicOrigins);
}
