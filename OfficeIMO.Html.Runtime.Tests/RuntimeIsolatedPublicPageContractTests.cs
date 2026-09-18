using System.Net;
using OfficeIMO.Html.Runtime.Rendering;
using Xunit;

namespace OfficeIMO.Html.Runtime.Tests;

public class RuntimeIsolatedPublicPageContractTests {
    [Fact]
    public void PublicPageRequestOwnsContentAndNetworkAuthority() {
        var request = new HtmlIsolatedPublicPageRequest {
            ScenarioId = "wpt-first-letter",
            Url = new Uri("https://wpt.live/reference.html"),
            SourceLicense = "BSD-3-Clause",
            Runtime = new HtmlScriptRequest {
                Profile = HtmlRuntimeProfile.WebApplicationV1,
                Html = "<p>caller content</p>"
            }
        };

        ArgumentException error = Assert.Throws<ArgumentException>(() => request.Validate());

        Assert.Contains("owned by the isolated public-page workflow", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PublicPageProfilePreservesEveryStricterCallerResourceLimit() {
        var caller = new HtmlRuntimeResourcePolicy {
            Timeout = TimeSpan.FromSeconds(2),
            MaxConcurrentRequests = 2,
            MaxRequests = 7,
            MaxResourceBytes = 1024,
            MaxTotalBytes = 4096,
            MaxRequestBytes = 128,
            MaxTotalRequestBytes = 256,
            MaxRedirects = 1
        }.Snapshot();

        HtmlRuntimeResourcePolicy effective = HtmlIsolatedPublicPageWorkflow.BoundedPolicy(
            caller, new[] { new Uri("https://assets.example/") });

        Assert.False(effective.AllowNetwork);
        Assert.Equal(caller.Timeout, effective.Timeout);
        Assert.Equal(caller.MaxConcurrentRequests, effective.MaxConcurrentRequests);
        Assert.Equal(caller.MaxRequests, effective.MaxRequests);
        Assert.Equal(caller.MaxResourceBytes, effective.MaxResourceBytes);
        Assert.Equal(caller.MaxTotalBytes, effective.MaxTotalBytes);
        Assert.Equal(caller.MaxRequestBytes, effective.MaxRequestBytes);
        Assert.Equal(caller.MaxTotalRequestBytes, effective.MaxTotalRequestBytes);
        Assert.Equal(caller.MaxRedirects, effective.MaxRedirects);
        Assert.Equal(new Uri("https://assets.example/"), Assert.Single(effective.AllowedOrigins));
    }

    [Fact]
    public void PublicPageRequestRejectsCallerSuppliedNetworkOrigins() {
        var request = new HtmlIsolatedPublicPageRequest {
            ScenarioId = "network-authority",
            Url = new Uri("https://example.com/"),
            SourceLicense = "fixture",
            Runtime = new HtmlScriptRequest {
                Profile = HtmlRuntimeProfile.WebApplicationV1,
                ResourcePolicy = new HtmlRuntimeResourcePolicy {
                    AllowedOrigins = new[] { new Uri("https://assets.example/") }
                }
            }
        };

        Assert.Throws<ArgumentException>(() => request.Validate());
    }

    [Fact]
    public void PublicPageRequestRejectsCallerSuppliedDynamicReplays() {
        var fetch = new HtmlRuntimeFetchRequest(new Uri("https://example.com/submit"), "POST", body: new byte[] { 1 });
        var request = new HtmlIsolatedPublicPageRequest {
            ScenarioId = "replay-authority",
            Url = new Uri("https://example.com/"),
            SourceLicense = "fixture",
            Runtime = new HtmlScriptRequest {
                Profile = HtmlRuntimeProfile.WebApplicationV1,
                FetchReplays = new[] {
                    new HtmlRuntimeFetchReplay(fetch, 1, HtmlRuntimeResource.FromText(fetch.Url, "accepted", "text/plain"))
                }
            }
        };

        ArgumentException error = Assert.Throws<ArgumentException>(() => request.Validate());
        Assert.Contains("dynamic replays", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DynamicReplayTranscriptRequiresExactOrderedConsumption() {
        var request = new HtmlRuntimeFetchRequest(new Uri("https://example.com/submit"), "POST", body: new byte[] { 1 });
        var first = new HtmlRuntimeFetchReplay(request, 1, HtmlRuntimeResource.FromText(request.Url, "first", "text/plain"));
        var second = new HtmlRuntimeFetchReplay(request, 2, HtmlRuntimeResource.FromText(request.Url, "second", "text/plain"));

        HtmlRuntimeFetchTranscript.Validate(new[] { first, second }, new[] { first.Identity, second.Identity });
        Assert.Throws<HtmlScriptRuntimeException>(() => HtmlRuntimeFetchTranscript.Validate(new[] { first }, Array.Empty<string>()));
        Assert.Throws<HtmlScriptRuntimeException>(() => HtmlRuntimeFetchTranscript.Validate(
            new[] { first, second }, new[] { second.Identity, first.Identity }));
    }

    [Fact]
    public void PublicPageRequestNormalizesExplicitDynamicMethodAuthority() {
        var methods = new List<string> { "get", "POST", "post" };
        var request = new HtmlIsolatedPublicPageRequest {
            ScenarioId = "dynamic-methods",
            Url = new Uri("https://example.com/"),
            SourceLicense = "fixture",
            AllowedDynamicRequestMethods = methods
        };

        HtmlIsolatedPublicPageRequest.Snapshot snapshot = request.Validate();
        methods[1] = "DELETE";

        Assert.Equal(new[] { "GET", "POST" }, snapshot.AllowedDynamicRequestMethods);
        Assert.Throws<ArgumentException>(() => new HtmlIsolatedPublicPageRequest {
            ScenarioId = "bad-method",
            Url = new Uri("https://example.com/"),
            SourceLicense = "fixture",
            AllowedDynamicRequestMethods = new[] { "TRACE" }
        }.Validate());
    }

    [Fact]
    public void ExecutionOptionsRequireImmutableImageAndSnapshotCommandPrefix() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-public-contract-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string renderer = Path.Combine(directory, "renderer.dll");
        string worker = Path.Combine(directory, "worker.dll");
        File.WriteAllBytes(renderer, [1]);
        File.WriteAllBytes(worker, [2]);
        var arguments = new List<string> { "-d", "Ubuntu", "--exec", "podman" };
        try {
            var options = new HtmlIsolatedPublicPageExecutionOptions {
                ImageId = "sha256:" + new string('A', 64),
                PodmanCommand = "wsl.exe",
                PodmanCommandArguments = arguments,
                PublishedRendererAssemblyPath = renderer,
                PublishedWorkerAssemblyPath = worker
            };

            HtmlIsolatedPublicPageExecutionOptions.Snapshot snapshot = options.Validate();
            arguments[1] = "changed";

            Assert.Equal("sha256:" + new string('a', 64), snapshot.ImageId);
            Assert.Equal(new[] { "-d", "Ubuntu", "--exec", "podman" }, snapshot.Arguments);
            Assert.Throws<ArgumentException>(() => new HtmlIsolatedPublicPageExecutionOptions {
                ImageId = "latest",
                PublishedRendererAssemblyPath = renderer,
                PublishedWorkerAssemblyPath = worker
            }.Validate());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void EvidenceRetainsProvenanceAndOnlyExplicitlyRetainedBytes() {
        var requested = new Uri("https://example.com/start");
        var final = new Uri("https://example.com/final");
        var connected = IPAddress.Parse("93.184.216.34");
        byte[] bytes = [1, 2, 3];
        var acquired = new HtmlPublicResourceResult(
            new HtmlRuntimeResource(requested, bytes, "text/html", finalUrl: final, redirectCount: 1),
            Array.AsReadOnly(new[] { new HtmlPublicRedirect(requested, final, 302, connected) }),
            DateTimeOffset.Parse("2026-09-18T12:00:00Z"), connected, "digest");

        var retained = new HtmlPublicResourceEvidence(acquired, retainBytes: true);
        var digestOnly = new HtmlPublicResourceEvidence(acquired, retainBytes: false);
        var output = new HtmlIsolatedPageOutput("screen.png", "image/png", bytes);
        bytes[0] = 9;

        Assert.Equal(final, retained.FinalUrl);
        Assert.Equal(connected, Assert.Single(retained.Redirects).ConnectedAddress);
        Assert.Equal(new byte[] { 1, 2, 3 }, retained.Content!.Value.ToArray());
        Assert.Null(digestOnly.Content);
        Assert.Equal(new byte[] { 1, 2, 3 }, output.Content.ToArray());
    }

    [Fact]
    public void DynamicEvidenceRetainsNamesAndDigestsWithoutHeaderValuesOrRequestBytes() {
        var url = new Uri("https://example.com/submit");
        var request = new HtmlRuntimeFetchRequest(url, "POST",
            new Dictionary<string, string> { ["X-Secret"] = "private", ["Content-Type"] = "text/plain" },
            System.Text.Encoding.UTF8.GetBytes("payload"));
        var discovery = new HtmlRuntimeFetchDiscovery(request, 2);
        var acquired = new HtmlPublicResourceResult(HtmlRuntimeResource.FromText(url, "accepted", "text/plain"),
            Array.Empty<HtmlPublicRedirect>(), DateTimeOffset.Parse("2026-09-18T12:00:00Z"),
            IPAddress.Parse("93.184.216.34"), "digest", discovery);

        var evidence = new HtmlPublicResourceEvidence(acquired, retainBytes: false);

        Assert.Equal("POST", evidence.RequestMethod);
        Assert.Equal(2, evidence.RequestOccurrence);
        Assert.Equal(new[] { "Content-Type", "X-Secret" }, evidence.RequestHeaderNames);
        Assert.Equal(7, evidence.RequestBodyByteCount);
        Assert.Equal(64, evidence.RequestBodySha256!.Length);
        Assert.DoesNotContain("private", string.Join(",", evidence.RequestHeaderNames), StringComparison.Ordinal);
        Assert.Null(evidence.Content);
    }

    [Fact]
    public void FailureEvidenceKeepsExpectedAndReportedPayloadDigestsSeparate() {
        var response = new HtmlPublicRenderResponse {
            RendererSha256 = "reported-renderer",
            WorkerSha256 = "reported-worker",
            RendererFilesSha256 = "reported-renderer-files",
            WorkerFilesSha256 = "reported-worker-files"
        };

        var evidence = new HtmlIsolatedPublicPageFailureEvidence(HtmlIsolatedPublicPagePhase.IsolatedStartup,
            Array.Empty<HtmlPublicResourceEvidence>(), Array.Empty<HtmlPublicSkippedResource>(),
            "sha256:" + new string('a', 64), "container", true, null, response,
            "expected-renderer", "expected-worker", "expected-renderer-files", "expected-worker-files");

        Assert.Equal("expected-renderer", evidence.ExpectedRendererSha256);
        Assert.Equal("reported-renderer", evidence.ReportedRendererSha256);
        Assert.Equal("expected-worker-files", evidence.ExpectedWorkerFilesSha256);
        Assert.Equal("reported-worker-files", evidence.ReportedWorkerFilesSha256);
    }

    [Fact]
    public void WorkerStageMapsToFailurePhase() {
        Assert.Equal(HtmlIsolatedPublicPagePhase.ResourceDiscovery,
            HtmlIsolatedPublicPageWorkflow.WorkerPhase(HtmlPublicRenderStage.ResourceDiscovery));
        Assert.Equal(HtmlIsolatedPublicPagePhase.Rendering,
            HtmlIsolatedPublicPageWorkflow.WorkerPhase(HtmlPublicRenderStage.Rendering));
        Assert.Equal(HtmlIsolatedPublicPagePhase.Output,
            HtmlIsolatedPublicPageWorkflow.WorkerPhase(HtmlPublicRenderStage.Output));
    }
}
