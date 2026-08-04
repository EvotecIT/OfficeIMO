using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentProducerCorpusTests {
    [Fact]
    public void ProducerManifest_Covers_All_Authored_Fixtures_With_Hashes_And_Advanced_Evidence() {
        string fixturePath = Path.Combine(AppContext.BaseDirectory, "Fixtures");
        ProducerManifest manifest = LoadManifest(fixturePath);
        string[] actual = Directory.GetFiles(fixturePath)
            .Where(path => new[] { ".odt", ".ods", ".odp" }.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            .Where(path => !string.Equals(Path.GetFileName(path), "extreme-repeats.ods", StringComparison.OrdinalIgnoreCase))
            .Select(Path.GetFileName)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray()!;
        string[] declared = manifest.Fixtures.Select(fixture => fixture.File)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(1, manifest.SchemaVersion);
        Assert.Equal(actual, declared);
        Assert.Equal(6, manifest.Fixtures.Count);
        Assert.Equal(6, manifest.Fixtures.Select(fixture => fixture.Producer).Distinct(StringComparer.Ordinal).Count());

        foreach (ProducerFixture fixture in manifest.Fixtures) {
            string path = Path.Combine(fixturePath, fixture.File);
            byte[] bytes = File.ReadAllBytes(path);
            Assert.Equal(fixture.Bytes, bytes.Length);
            Assert.Equal(fixture.Sha256, ComputeSha256(bytes));
            Assert.NotEmpty(fixture.Evidence);

            OdfDocument document = OdfDocument.Load(path);
            Assert.Equal(fixture.Kind, document.Kind switch {
                OdfDocumentKind.Text => "ODT",
                OdfDocumentKind.Spreadsheet => "ODS",
                OdfDocumentKind.Presentation => "ODP",
                _ => throw new InvalidOperationException("Unexpected OpenDocument kind.")
            });
            Assert.True(document.Validate().IsValid);
        }

        Assert.Equal(
            new[] { "drawings", "embedded-content", "formulas", "styles", "unknown-package-content" },
            manifest.CapabilityEvidence.Select(evidence => evidence.Id).OrderBy(id => id, StringComparer.Ordinal));
        Assert.All(manifest.CapabilityEvidence, evidence => {
            Assert.False(string.IsNullOrWhiteSpace(evidence.Contract));
            Assert.Contains('.', evidence.Test);
        });
    }

    [Fact]
    public void ExternalProducerEvidence_Is_HashPinned_Without_Claiming_Redistribution() {
        ProducerManifest manifest = LoadManifest(Path.Combine(AppContext.BaseDirectory, "Fixtures"));
        ExternalArtifact google = Assert.Single(manifest.ExternalArtifacts);

        Assert.Equal("Google Docs", google.Producer);
        Assert.Equal("ODT", google.Kind);
        Assert.StartsWith("https://docs.google.com/", google.SourceUrl, StringComparison.Ordinal);
        Assert.Equal(64, google.SemanticTextSha256.Length);
        Assert.True(google.MinBytes > 0);
        Assert.True(google.MaxBytes > google.MinBytes);
        Assert.True(google.ParagraphCount > 0);
        Assert.False(google.RedistributionAllowed);
        Assert.False(string.IsNullOrWhiteSpace(google.Note));
    }

    private static ProducerManifest LoadManifest(string fixturePath) {
        using FileStream stream = File.OpenRead(Path.Combine(fixturePath, "producer-manifest.json"));
        var serializer = new DataContractJsonSerializer(typeof(ProducerManifest));
        return Assert.IsType<ProducerManifest>(serializer.ReadObject(stream));
    }

    private static string ComputeSha256(byte[] bytes) {
        using SHA256 sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
    }

    [DataContract]
    private sealed class ProducerManifest {
        [DataMember(Name = "schemaVersion")]
        public int SchemaVersion { get; set; }
        [DataMember(Name = "fixtures")]
        public List<ProducerFixture> Fixtures { get; set; } = new();
        [DataMember(Name = "externalArtifacts")]
        public List<ExternalArtifact> ExternalArtifacts { get; set; } = new();
        [DataMember(Name = "capabilityEvidence")]
        public List<CapabilityEvidence> CapabilityEvidence { get; set; } = new();
    }

    [DataContract]
    private sealed class ProducerFixture {
        [DataMember(Name = "file")]
        public string File { get; set; } = string.Empty;
        [DataMember(Name = "kind")]
        public string Kind { get; set; } = string.Empty;
        [DataMember(Name = "producer")]
        public string Producer { get; set; } = string.Empty;
        [DataMember(Name = "producerVersion")]
        public string ProducerVersion { get; set; } = string.Empty;
        [DataMember(Name = "producedOn")]
        public string ProducedOn { get; set; } = string.Empty;
        [DataMember(Name = "bytes")]
        public int Bytes { get; set; }
        [DataMember(Name = "sha256")]
        public string Sha256 { get; set; } = string.Empty;
        [DataMember(Name = "evidence")]
        public List<string> Evidence { get; set; } = new();
    }

    [DataContract]
    private sealed class ExternalArtifact {
        [DataMember(Name = "id")]
        public string Id { get; set; } = string.Empty;
        [DataMember(Name = "kind")]
        public string Kind { get; set; } = string.Empty;
        [DataMember(Name = "producer")]
        public string Producer { get; set; } = string.Empty;
        [DataMember(Name = "producerVersion")]
        public string ProducerVersion { get; set; } = string.Empty;
        [DataMember(Name = "sourceUrl")]
        public string SourceUrl { get; set; } = string.Empty;
        [DataMember(Name = "observedAtUtc")]
        public string ObservedAtUtc { get; set; } = string.Empty;
        [DataMember(Name = "minBytes")]
        public int MinBytes { get; set; }
        [DataMember(Name = "maxBytes")]
        public int MaxBytes { get; set; }
        [DataMember(Name = "paragraphCount")]
        public int ParagraphCount { get; set; }
        [DataMember(Name = "semanticTextSha256")]
        public string SemanticTextSha256 { get; set; } = string.Empty;
        [DataMember(Name = "redistributionAllowed")]
        public bool RedistributionAllowed { get; set; }
        [DataMember(Name = "note")]
        public string Note { get; set; } = string.Empty;
    }

    [DataContract]
    private sealed class CapabilityEvidence {
        [DataMember(Name = "id")]
        public string Id { get; set; } = string.Empty;
        [DataMember(Name = "contract")]
        public string Contract { get; set; } = string.Empty;
        [DataMember(Name = "test")]
        public string Test { get; set; } = string.Empty;
    }
}
