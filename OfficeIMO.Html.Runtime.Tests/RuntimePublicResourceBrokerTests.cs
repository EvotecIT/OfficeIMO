using System.Net;
using OfficeIMO.Html.Runtime;
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
}
