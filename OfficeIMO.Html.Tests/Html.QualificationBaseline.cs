using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Qualification;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlQualificationBaselineTests {
    [Fact]
    public void FrozenCorpus_LoadsWithDeclaredInputsAndProfiles() {
        HtmlQualificationCorpus corpus = HtmlQualificationCorpus.Load();

        Assert.Equal("officeimo-html-h0-static-v1", corpus.Manifest.CorpusId);
        Assert.Equal(7, corpus.Manifest.Files.Count);
        Assert.Equal(new[] { "wide", "narrow", "print-a4" }, corpus.Manifest.Profiles.Select(profile => profile.Id));
        Assert.Equal(64, corpus.ManifestSha256.Length);
        Assert.Contains("Northstar Operations", corpus.EntryHtml);
    }

    [Fact]
    public async Task FrozenCorpus_RendersAllProfilesAndResolvesEveryDeclaredResource() {
        HtmlQualificationCorpus corpus = HtmlQualificationCorpus.Load();
        string[] expectedResources = corpus.Manifest.Files
            .Where(file => file.Role is "stylesheet" or "font" or "image")
            .Select(file => file.Path)
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();

        foreach (HtmlQualificationProfile profile in corpus.Manifest.Profiles) {
            var resolved = new HashSet<string>(StringComparer.Ordinal);
            var options = new HtmlRenderOptions {
                Mode = profile.Mode == "paged" ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous,
                ViewportWidth = profile.ViewportWidth,
                ViewportHeight = profile.ViewportHeight,
                PageSize = OfficePageSizes.A4,
                Margins = HtmlRenderMargins.All(40D),
                BaseUri = corpus.BaseUri,
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                ResourceResolver = (request, cancellationToken) => {
                    cancellationToken.ThrowIfCancellationRequested();
                    Assert.True(corpus.TryResolve(request.Uri, out HtmlQualificationInputFile file, out byte[] bytes),
                        "The renderer requested an undeclared resource: " + request.Uri.AbsoluteUri);
                    resolved.Add(file.Path);
                    return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(bytes, file.MediaType, request.Uri, 0));
                }
            };

            HtmlConversionDocument source = HtmlConversionDocument.Parse(corpus.EntryHtml, new HtmlConversionDocumentOptions {
                BaseUri = corpus.BaseUri,
                Trust = HtmlInputTrust.Untrusted,
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });
            HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(source, options, CancellationToken.None);

            Assert.Equal(profile.ExpectedPageCount, rendered.Pages.Count);
            Assert.False(rendered.HasLoss, string.Join(Environment.NewLine, rendered.Diagnostics.Select(diagnostic =>
                diagnostic.Code + ": " + diagnostic.Message)));
            foreach (string marker in corpus.Manifest.TextMarkers) Assert.Contains(marker, rendered.Text);
            Assert.Equal(expectedResources, resolved.OrderBy(path => path, StringComparer.Ordinal));
        }
    }

    [Fact]
    public void FrozenCorpus_RejectsChangedInputBytes() {
        string sourceRoot = HtmlQualificationCorpus.ResolveDefaultRoot();
        string temporaryRoot = Path.Combine(Path.GetTempPath(), "OfficeIMO.Html.Qualification.Tests." + Guid.NewGuid().ToString("N"));
        try {
            CopyTree(sourceRoot, temporaryRoot);
            string htmlPath = Path.Combine(temporaryRoot, "representative-report", "index.html");
            File.AppendAllText(htmlPath, " ");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => HtmlQualificationCorpus.Load(temporaryRoot));
            Assert.Contains("representative-report/index.html", exception.Message);
        } finally {
            if (Directory.Exists(temporaryRoot)) Directory.Delete(temporaryRoot, recursive: true);
        }
    }

    [Fact]
    public void FrozenCorpus_RejectsProfileIdsThatCanEscapeOutputRoots() {
        string sourceRoot = HtmlQualificationCorpus.ResolveDefaultRoot();
        string temporaryRoot = Path.Combine(Path.GetTempPath(), "OfficeIMO.Html.Qualification.Tests." + Guid.NewGuid().ToString("N"));
        try {
            CopyTree(sourceRoot, temporaryRoot);
            string manifestPath = Path.Combine(temporaryRoot, "manifest.json");
            string manifest = File.ReadAllText(manifestPath).Replace("\"id\": \"wide\"", "\"id\": \"../escape\"");
            File.WriteAllText(manifestPath, manifest);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => HtmlQualificationCorpus.Load(temporaryRoot));
            Assert.Contains("one safe path segment", exception.Message);
        } finally {
            if (Directory.Exists(temporaryRoot)) Directory.Delete(temporaryRoot, recursive: true);
        }
    }

    private static void CopyTree(string source, string destination) {
        Directory.CreateDirectory(destination);
        foreach (string directory in Directory.GetDirectories(source, "*", SearchOption.AllDirectories)) {
            Directory.CreateDirectory(Path.Combine(destination, RelativePath(source, directory)));
        }
        foreach (string file in Directory.GetFiles(source, "*", SearchOption.AllDirectories)) {
            string target = Path.Combine(destination, RelativePath(source, file));
            File.Copy(file, target);
        }
    }

    private static string RelativePath(string root, string path) {
        var rootUri = new Uri(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar);
        return Uri.UnescapeDataString(rootUri.MakeRelativeUri(new Uri(path)).ToString()).Replace('/', Path.DirectorySeparatorChar);
    }
}
