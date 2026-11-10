using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The documented edge rules of merging that the corpus pairs do not reach: null removal, kind conflicts, anchor
// collisions, merge-by-key, and overlays that land on aliases.
public class YamlDocumentMergeTests {
    private static string Merge(string baseText, string overlayText, YamlMergeOptions? options, out YamlMergeReport report) {
        YamlDocument document = YamlDocument.Parse(baseText);
        report = document.Merge(YamlDocument.Parse(overlayText), options);
        return Save(document);
    }

    private static string Save(YamlDocument document) {
        var stream = new MemoryStream();
        document.Save(stream);
        return new System.Text.UTF8Encoding(false).GetString(stream.ToArray());
    }

    [Fact]
    public void NullRemovesTheBaseKeyOnlyWhenAsked() {
        const string baseText = "image:\n  tag: 1.0 # pinned\nresources:\n  limits:\n    cpu: 1\nreplicas: 2\n";
        const string overlay = "resources: null\nmissing: ~\n";

        string removed = Merge(baseText, overlay, new YamlMergeOptions { NullRemovesKey = true }, out YamlMergeReport report);
        Assert.Equal("image:\n  tag: 1.0 # pinned\nreplicas: 2\n", removed);
        Assert.Equal(new[] { "resources" }, report.Removed);
        Assert.Empty(report.Added);

        string replaced = Merge(baseText, overlay, null, out report);
        Assert.Equal("image:\n  tag: 1.0 # pinned\nresources: null\nreplicas: 2\nmissing: ~\n", replaced);
        Assert.Equal(new[] { "resources" }, report.Replaced);
        Assert.Equal(new[] { "missing" }, report.Added);
        Assert.Empty(report.Diagnostics);
    }

    [Fact]
    public void KindConflictsLetTheOverlayWinAndAreReported() {
        string merged = Merge("service:\n  port: 80\nname: web\n", "service: disabled\nname:\n  first: web\n", null, out YamlMergeReport report);

        Assert.Equal("service: disabled\nname:\n  first: web\n", merged);
        Assert.Equal(new[] { "service", "name" }, report.Replaced);
        Assert.Equal(2, report.Diagnostics.Count);
        Assert.All(report.Diagnostics, d => Assert.Equal(OfficeDiagnosticCodes.YamlMergeKindConflict, d.Code));
        Assert.Equal("2:3", report.Diagnostics[0].Location);
    }

    [Fact]
    public void OverlayAnchorsAreCopiedAndRenamedWhenTheBaseUsesTheName() {
        const string baseText = "defaults: &common\n  retries: 3\njob:\n  <<: *common\n";
        const string overlay = "extra: &common\n  timeout: 30\nother: *common\n";

        YamlDocument document = YamlDocument.Parse(baseText);
        YamlMergeReport report = document.Merge(YamlDocument.Parse(overlay));
        string merged = Save(document);

        Assert.Equal(baseText + "extra: &common-2\n  timeout: 30\nother: *common-2\n", merged);
        OfficeDiagnostic renamed = Assert.Single(report.Diagnostics);
        Assert.Equal(OfficeDiagnosticCodes.YamlMergeAnchorRenamed, renamed.Code);
        Assert.Equal(30, YamlDocument.Parse(merged).Select("other.timeout")!.GetValue<int>());
    }

    [Fact]
    public void MergeByKeyMergesMatchingItemsAndAppendsTheRest() {
        const string baseText = "containers:\n- name: app\n  image: app:1\n  ports:\n  - 80\n- name: sidecar # logs\n  image: log:1\n";
        const string overlay = "containers:\n- name: sidecar\n  image: log:2\n- name: metrics\n  image: prom:1\n";

        string merged = Merge(baseText, overlay, new YamlMergeOptions { Sequences = YamlSequenceMerge.MergeByKey("name") }, out YamlMergeReport report);

        Assert.Equal("containers:\n- name: app\n  image: app:1\n  ports:\n  - 80\n- name: sidecar # logs\n  image: log:2\n- name: metrics\n  image: prom:1\n", merged);
        Assert.Equal(new[] { "containers[1].image" }, report.Replaced);
        Assert.Equal(new[] { "containers[2]" }, report.Added);
    }

    [Fact]
    public void OverlayValuesThatLandOnAnAliasEditTheAnchoredNode() {
        const string baseText = "base: &limits\n  cpu: 1\nworker: *limits\n";

        string merged = Merge(baseText, "worker:\n  cpu: 2\n", null, out YamlMergeReport report);

        Assert.Equal("base: &limits\n  cpu: 2\nworker: *limits\n", merged);
        Assert.Equal(new[] { "worker.cpu" }, report.Replaced);
        Assert.Equal(OfficeDiagnosticCodes.YamlAliasedNodeEdited, Assert.Single(report.Diagnostics).Code);
    }

    [Fact]
    public void MergeKeysFollowTheTypedViewAndAreNeverFlattened() {
        const string baseText = "%YAML 1.1\n---\ndefaults: &defaults\n  retries: 3\n  timeout: 10\njob:\n  <<: *defaults\n  name: build\n";
        YamlDocument document = YamlDocument.Parse(baseText);
        YamlMapping overlay = YamlDocument.Parse("timeout: 20\nretries: 3\n").Root!.AsMapping();

        YamlMergeReport report = document.Select("job")!.AsMapping().Merge(overlay);

        Assert.Equal(baseText + "  timeout: 20\n", Save(document));
        Assert.Equal(new[] { "job.timeout" }, report.Replaced);
        Assert.Empty(report.Added);
        Assert.Equal(10, document.Select("defaults.timeout")!.GetValue<int>());
    }
}
