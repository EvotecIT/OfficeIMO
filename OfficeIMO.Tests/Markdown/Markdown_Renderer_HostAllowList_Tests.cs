using System.Collections.Generic;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public class Markdown_Renderer_HostAllowList_Tests {
        [Fact]
        public void HostAllowList_ExactMatch() {
            Assert.True(HostAllowList.IsAllowed("example.com", new List<string> { "example.com" }));
            Assert.False(HostAllowList.IsAllowed("sub.example.com", new List<string> { "example.com" }));
        }

        [Fact]
        public void HostAllowList_ApexAndSubdomains_DotPattern() {
            Assert.True(HostAllowList.IsAllowed("example.com", new List<string> { ".example.com" }));
            Assert.True(HostAllowList.IsAllowed("a.example.com", new List<string> { ".example.com" }));
            Assert.True(HostAllowList.IsAllowed("a.b.example.com", new List<string> { ".example.com" }));
            Assert.False(HostAllowList.IsAllowed("badexample.com", new List<string> { ".example.com" }));
        }

        [Fact]
        public void HostAllowList_SubdomainsOnly_StarDotPattern() {
            Assert.False(HostAllowList.IsAllowed("example.com", new List<string> { "*.example.com" }));
            Assert.True(HostAllowList.IsAllowed("a.example.com", new List<string> { "*.example.com" }));
            Assert.True(HostAllowList.IsAllowed("a.b.example.com", new List<string> { "*.example.com" }));
            Assert.False(HostAllowList.IsAllowed("badexample.com", new List<string> { "*.example.com" }));
        }

        [Fact]
        public void HostAllowList_NormalizesCaseAndTrailingDot() {
            Assert.True(HostAllowList.IsAllowed("Example.COM.", new List<string> { "example.com" }));
            Assert.True(HostAllowList.IsAllowed("a.Example.COM.", new List<string> { ".example.com." }));
            Assert.True(HostAllowList.IsAllowed("a.example.com", new List<string> { "*.example.com." }));
        }

        [Fact]
        public void HostAllowList_AllowsWildcardButNotEmptyHost() {
            Assert.True(HostAllowList.IsAllowed("example.com", new List<string> { "*" }));
            Assert.False(HostAllowList.IsAllowed("", new List<string> { "*" }));
            Assert.False(HostAllowList.IsAllowed(null, new List<string> { "*" }));
        }

        [Fact]
        public void HostAllowList_IgnoresPortAndSchemeInPatterns() {
            Assert.True(HostAllowList.IsAllowed("example.com", new List<string> { "example.com:443" }));
            Assert.True(HostAllowList.IsAllowed("example.com", new List<string> { "https://example.com" }));
            Assert.True(HostAllowList.IsAllowed("a.example.com", new List<string> { ".example.com:8443" }));
            Assert.True(HostAllowList.IsAllowed("a.example.com", new List<string> { "*.example.com:8443" }));
        }

        [Fact]
        public void HostAllowList_BracketedIpv6Patterns() {
            Assert.True(HostAllowList.IsAllowed("fe80::1", new List<string> { "[fe80::1]" }));
            Assert.True(HostAllowList.IsAllowed("fe80::1", new List<string> { "[fe80::1]:443" }));
            Assert.False(HostAllowList.IsAllowed("fe80::2", new List<string> { "[fe80::1]" }));
        }

        [Fact]
        public void UrlOriginPolicy_RespectsHostAllowList_ForAbsoluteHttpLinksAndImages() {
            var o = new HtmlOptions();
            o.AllowedHttpLinkHosts.Add("example.com");
            o.AllowedHttpImageHosts.Add(".img.example.com");

            Assert.True(UrlOriginPolicy.IsAllowedHttpLink(o, "https://example.com/a"));
            Assert.False(UrlOriginPolicy.IsAllowedHttpLink(o, "https://evil.com/a"));

            Assert.True(UrlOriginPolicy.IsAllowedHttpImage(o, "https://img.example.com/a.png"));
            Assert.True(UrlOriginPolicy.IsAllowedHttpImage(o, "https://a.img.example.com/a.png"));
            Assert.False(UrlOriginPolicy.IsAllowedHttpImage(o, "https://example.com/a.png"));
        }

        [Fact]
        public void UrlOriginPolicy_HostAllowList_DoesNotBlockRelativeUrls() {
            var o = new HtmlOptions { BaseUri = new System.Uri("https://base.example/") };
            o.AllowedHttpLinkHosts.Add("example.com");
            o.AllowedHttpImageHosts.Add("example.com");

            // Allowlists apply to absolute HTTP(S) only; relative stays allowed.
            Assert.True(UrlOriginPolicy.IsAllowedHttpLink(o, "/docs"));
            Assert.True(UrlOriginPolicy.IsAllowedHttpImage(o, "/img.png"));
        }

        [Fact]
        public void UrlOriginPolicy_HostAllowList_AppliesToProtocolRelativeUrls() {
            var o = new HtmlOptions { BaseUri = new System.Uri("https://base.example/") };
            o.AllowedHttpLinkHosts.Add("example.com");

            Assert.True(UrlOriginPolicy.IsAllowedHttpLink(o, "//example.com/a"));
            Assert.False(UrlOriginPolicy.IsAllowedHttpLink(o, "//evil.com/a"));
        }
    }
}

