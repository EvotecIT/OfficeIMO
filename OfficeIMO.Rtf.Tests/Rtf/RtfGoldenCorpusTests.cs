using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfGoldenCorpusTests {
    [Fact]
    public void CorpusManifest_Covers_Every_Fixture_With_Provenance_And_Stable_Hash() {
        string corpusPath = GetCorpusPath();
        CorpusManifest manifest = LoadManifest(corpusPath);
        string[] actualFiles = Directory.GetFiles(corpusPath, "*.rtf", SearchOption.AllDirectories)
            .Select(path => ToManifestPath(corpusPath, path))
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        string[] declaredFiles = manifest.Fixtures.Select(fixture => fixture.File)
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(1, manifest.SchemaVersion);
        Assert.Equal(actualFiles, declaredFiles);
        Assert.Equal(manifest.Fixtures.Count, manifest.Fixtures.Select(fixture => fixture.Id).Distinct(StringComparer.Ordinal).Count());

        foreach (CorpusFixture fixture in manifest.Fixtures) {
            string path = GetFixturePath(corpusPath, fixture.File);
            Assert.True(File.Exists(path), fixture.File);
            Assert.False(string.IsNullOrWhiteSpace(fixture.Producer));
            Assert.False(string.IsNullOrWhiteSpace(fixture.ProducerVersion));
            Assert.False(string.IsNullOrWhiteSpace(fixture.Origin));
            Assert.False(string.IsNullOrWhiteSpace(fixture.License));
            Assert.True(fixture.RedistributionAllowed, fixture.File);
            Assert.Equal(fixture.Sha256, ComputeSha256(path));
            Assert.NotEmpty(fixture.Features);
            Assert.Contains("Core", fixture.Adapters);
            if (fixture.EvidenceClass == "upstream-regression") {
                Assert.StartsWith("https://github.com/LibreOffice/core/blob/", fixture.SourceUrl, StringComparison.Ordinal);
                Assert.Equal("MPL-2.0", fixture.License);
            }
        }
    }

    [Fact]
    public void CorpusFixtures_Preserve_Bytes_Reparse_Normalized_Output_And_Meet_Semantic_Expectations() {
        string corpusPath = GetCorpusPath();
        CorpusManifest manifest = LoadManifest(corpusPath);

        foreach (CorpusFixture fixture in manifest.Fixtures) {
            string path = GetFixturePath(corpusPath, fixture.File);
            byte[] sourceBytes = File.ReadAllBytes(path);
            RtfReadResult result = RtfDocument.Load(sourceBytes);

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == RtfDiagnosticSeverity.Error);
            Assert.Equal(sourceBytes, result.ToBytesLossless());
            string normalized = result.Document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
            RtfReadResult normalizedResult = RtfDocument.Read(normalized);
            Assert.DoesNotContain(normalizedResult.Diagnostics, diagnostic => diagnostic.Severity == RtfDiagnosticSeverity.Error);

            string source = result.ToRtfLossless();
            foreach (string control in fixture.RequiredControls) {
                Assert.Contains(control, source, StringComparison.Ordinal);
            }

            string semanticText = GetSemanticText(result.Document);
            foreach (string requiredText in fixture.RequiredText) {
                Assert.Contains(requiredText, semanticText, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void CorpusManifest_Adapter_Claims_Are_Executable() {
        string corpusPath = GetCorpusPath();
        CorpusManifest manifest = LoadManifest(corpusPath);

        foreach (CorpusFixture fixture in manifest.Fixtures.Where(item => item.Adapters.Count > 1)) {
            RtfDocument document = RtfDocument.Load(GetFixturePath(corpusPath, fixture.File)).Document;
            if (fixture.Adapters.Contains("Html")) {
                Assert.False(string.IsNullOrWhiteSpace(document.ToHtml()), fixture.Id + ": Html");
            }

            if (fixture.Adapters.Contains("Markdown")) {
                Assert.False(string.IsNullOrWhiteSpace(document.ToMarkdown()), fixture.Id + ": Markdown");
            }

            if (fixture.Adapters.Contains("Pdf")) {
                Assert.NotEmpty(document.ToPdf());
            }

            if (fixture.Adapters.Contains("Word")) {
                using WordDocument word = document.ToWordDocumentResult().Value;
                RtfDocument roundTrip = word.ToRtfDocument();
                Assert.False(string.IsNullOrWhiteSpace(roundTrip.ToRtf()), fixture.Id + ": Word");
            }
        }
    }

    [Fact]
    public void ProducerScorecard_Does_Not_Overstate_Unverified_Evidence() {
        CorpusManifest manifest = LoadManifest(GetCorpusPath());
        Dictionary<string, CorpusFixture> fixtures = manifest.Fixtures.ToDictionary(item => item.Id, StringComparer.Ordinal);

        Assert.Contains(manifest.ProducerCoverage, item => item.Producer == "Microsoft Word" && item.Status == "verified");
        Assert.Contains(manifest.ProducerCoverage, item => item.Producer == "Google Docs" && item.Status == "unverified");
        Assert.Contains(manifest.ProducerCoverage, item => item.Producer == "macOS TextEdit" && item.Status == "unverified");

        foreach (ProducerCoverage coverage in manifest.ProducerCoverage) {
            foreach (string fixtureId in coverage.FixtureIds) Assert.True(fixtures.ContainsKey(fixtureId), fixtureId);
            if (coverage.Status == "verified") {
                Assert.NotEmpty(coverage.FixtureIds);
                Assert.All(coverage.FixtureIds, id => Assert.Equal("producer-generated", fixtures[id].EvidenceClass));
                Assert.Contains(manifest.ReopenEvidence, evidence => evidence.FixtureId == coverage.FixtureIds[0] && evidence.Opened && evidence.RequiredTextPresent);
            }
        }
    }

    private static string GetSemanticText(RtfDocument document) {
        var parts = new List<string>();
        foreach (IRtfBlock block in document.Blocks) AppendBlockText(block, parts);
        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) parts.Add(headerFooter.ToPlainText());
        foreach (RtfNote note in document.Notes) parts.Add(note.ToPlainText());
        return string.Join("\n", parts);
    }

    private static void AppendBlockText(IRtfBlock block, List<string> parts) {
        switch (block) {
            case RtfParagraph paragraph:
                parts.Add(paragraph.ToPlainText());
                break;
            case RtfTable table:
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        foreach (IRtfBlock child in cell.Blocks) AppendBlockText(child, parts);
                    }
                }
                break;
            case RtfObject rtfObject:
                parts.Add(rtfObject.ToPlainText());
                break;
            case RtfShape shape:
                parts.Add(shape.ToPlainText());
                break;
        }
    }

    private static CorpusManifest LoadManifest(string corpusPath) {
        string path = Path.Combine(corpusPath, "corpus-manifest.json");
        using FileStream stream = File.OpenRead(path);
        var serializer = new DataContractJsonSerializer(typeof(CorpusManifest));
        return Assert.IsType<CorpusManifest>(serializer.ReadObject(stream));
    }

    private static string GetCorpusPath() => Path.Combine(AppContext.BaseDirectory, "Documents", "RtfCorpus");

    private static string GetFixturePath(string corpusPath, string manifestPath) =>
        Path.Combine(corpusPath, manifestPath.Replace('/', Path.DirectorySeparatorChar));

    private static string ToManifestPath(string corpusPath, string path) =>
        path.Substring(corpusPath.Length + 1).Replace(Path.DirectorySeparatorChar, '/');

    private static string ComputeSha256(string path) {
        using SHA256 sha = SHA256.Create();
        using FileStream stream = File.OpenRead(path);
        return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", string.Empty).ToLowerInvariant();
    }

    [DataContract]
    private sealed class CorpusManifest {
        [DataMember(Name = "schemaVersion")]
        public int SchemaVersion { get; set; }

        [DataMember(Name = "fixtures")]
        public List<CorpusFixture> Fixtures { get; set; } = new List<CorpusFixture>();

        [DataMember(Name = "reopenEvidence")]
        public List<ReopenEvidence> ReopenEvidence { get; set; } = new List<ReopenEvidence>();

        [DataMember(Name = "producerCoverage")]
        public List<ProducerCoverage> ProducerCoverage { get; set; } = new List<ProducerCoverage>();
    }

    [DataContract]
    private sealed class CorpusFixture {
        [DataMember(Name = "id")]
        public string Id { get; set; } = string.Empty;
        [DataMember(Name = "file")]
        public string File { get; set; } = string.Empty;
        [DataMember(Name = "evidenceClass")]
        public string EvidenceClass { get; set; } = string.Empty;
        [DataMember(Name = "producer")]
        public string Producer { get; set; } = string.Empty;
        [DataMember(Name = "producerVersion")]
        public string ProducerVersion { get; set; } = string.Empty;
        [DataMember(Name = "origin")]
        public string Origin { get; set; } = string.Empty;
        [DataMember(Name = "sourceUrl", EmitDefaultValue = false)]
        public string? SourceUrl { get; set; }
        [DataMember(Name = "license")]
        public string License { get; set; } = string.Empty;
        [DataMember(Name = "redistributionAllowed")]
        public bool RedistributionAllowed { get; set; }
        [DataMember(Name = "sha256")]
        public string Sha256 { get; set; } = string.Empty;
        [DataMember(Name = "features")]
        public List<string> Features { get; set; } = new List<string>();
        [DataMember(Name = "requiredText")]
        public List<string> RequiredText { get; set; } = new List<string>();
        [DataMember(Name = "requiredControls")]
        public List<string> RequiredControls { get; set; } = new List<string>();
        [DataMember(Name = "adapters")]
        public List<string> Adapters { get; set; } = new List<string>();
    }

    [DataContract]
    private sealed class ProducerCoverage {
        [DataMember(Name = "producer")]
        public string Producer { get; set; } = string.Empty;
        [DataMember(Name = "status")]
        public string Status { get; set; } = string.Empty;
        [DataMember(Name = "fixtureIds")]
        public List<string> FixtureIds { get; set; } = new List<string>();
        [DataMember(Name = "note")]
        public string Note { get; set; } = string.Empty;
    }

    [DataContract]
    private sealed class ReopenEvidence {
        [DataMember(Name = "fixtureId")]
        public string FixtureId { get; set; } = string.Empty;
        [DataMember(Name = "target")]
        public string Target { get; set; } = string.Empty;
        [DataMember(Name = "targetVersion")]
        public string TargetVersion { get; set; } = string.Empty;
        [DataMember(Name = "targetBuild")]
        public string TargetBuild { get; set; } = string.Empty;
        [DataMember(Name = "operation")]
        public string Operation { get; set; } = string.Empty;
        [DataMember(Name = "observedAtUtc")]
        public string ObservedAtUtc { get; set; } = string.Empty;
        [DataMember(Name = "opened")]
        public bool Opened { get; set; }
        [DataMember(Name = "paragraphCount")]
        public int ParagraphCount { get; set; }
        [DataMember(Name = "topLevelTableCount")]
        public int TopLevelTableCount { get; set; }
        [DataMember(Name = "nestedTableCount")]
        public int NestedTableCount { get; set; }
        [DataMember(Name = "requiredTextPresent")]
        public bool RequiredTextPresent { get; set; }
        [DataMember(Name = "outputBytes")]
        public int OutputBytes { get; set; }
        [DataMember(Name = "outputSha256")]
        public string OutputSha256 { get; set; } = string.Empty;
    }
}
