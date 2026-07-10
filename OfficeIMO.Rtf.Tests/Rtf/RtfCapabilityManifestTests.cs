using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfCapabilityManifestTests {
    [Fact]
    public void Capability_Manifest_Is_Complete_And_Represented_In_The_Living_Matrix() {
        string directory = Path.Combine(AppContext.BaseDirectory, "Documents", "RtfCapabilities");
        CapabilityManifest manifest;
        using (FileStream stream = File.OpenRead(Path.Combine(directory, "officeimo.rtf-capabilities.json"))) {
            var serializer = new DataContractJsonSerializer(typeof(CapabilityManifest));
            manifest = Assert.IsType<CapabilityManifest>(serializer.ReadObject(stream));
        }

        string matrix = File.ReadAllText(Path.Combine(directory, "officeimo.rtf-support-matrix.md"));
        string[] statuses = { "Full", "Broad", "Preserved", "Extractive" };
        string[] phases = { "P0", "P1", "P2" };
        string[] conversionClasses = { "Semantic", "Lossless", "Diagnostic", "Visual", "Extractive" };
        Assert.Equal(1, manifest.SchemaVersion);
        Assert.True(DateTime.TryParseExact(
            manifest.ReviewedOn,
            "yyyy-MM-dd",
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None,
            out _));
        Assert.Equal(manifest.Capabilities.Count, manifest.Capabilities.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.All(phases, phase => Assert.Contains(manifest.Capabilities, item => item.Phase == phase));

        foreach (Capability item in manifest.Capabilities) {
            Assert.False(string.IsNullOrWhiteSpace(item.Id));
            Assert.Contains(item.Phase, phases);
            Assert.Contains(item.Status, statuses);
            Assert.Contains(item.ConversionClass, conversionClasses);
            Assert.False(string.IsNullOrWhiteSpace(item.Owner));
            Assert.NotEmpty(item.PublicApi);
            Assert.NotEmpty(item.Evidence);
            Assert.All(item.PublicApi, api => Assert.False(string.IsNullOrWhiteSpace(api)));
            Assert.All(item.Evidence, evidence => Assert.False(string.IsNullOrWhiteSpace(evidence)));
            Assert.False(string.IsNullOrWhiteSpace(item.Boundary));
            Assert.Contains("<!-- capability:" + item.Id + " -->", matrix, StringComparison.Ordinal);
        }
    }

    [DataContract]
    private sealed class CapabilityManifest {
        [DataMember(Name = "schemaVersion")]
        public int SchemaVersion { get; set; }

        [DataMember(Name = "reviewedOn")]
        public string ReviewedOn { get; set; } = string.Empty;

        [DataMember(Name = "capabilities")]
        public List<Capability> Capabilities { get; set; } = new List<Capability>();
    }

    [DataContract]
    private sealed class Capability {
        [DataMember(Name = "id")]
        public string Id { get; set; } = string.Empty;

        [DataMember(Name = "phase")]
        public string Phase { get; set; } = string.Empty;

        [DataMember(Name = "owner")]
        public string Owner { get; set; } = string.Empty;

        [DataMember(Name = "status")]
        public string Status { get; set; } = string.Empty;

        [DataMember(Name = "conversionClass")]
        public string ConversionClass { get; set; } = string.Empty;

        [DataMember(Name = "publicApi")]
        public List<string> PublicApi { get; set; } = new List<string>();

        [DataMember(Name = "evidence")]
        public List<string> Evidence { get; set; } = new List<string>();

        [DataMember(Name = "boundary")]
        public string Boundary { get; set; } = string.Empty;
    }
}
