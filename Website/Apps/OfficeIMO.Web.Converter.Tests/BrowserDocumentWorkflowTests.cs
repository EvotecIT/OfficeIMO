using OfficeIMO.Provenance;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserDocumentWorkflowTests {
    [Fact]
    public void ProvenanceBrowserSelectionIsGeneratedFromTheSharedQualifiedCatalog() {
        Assert.Equal(OfficeProvenanceBufferWorkflow.SupportedExtensions, OfficeProvenanceWorkflowCatalog.BrowserExtensions);
        Assert.All(OfficeProvenanceWorkflowCatalog.BrowserCapabilities, capability => {
            Assert.True(capability.BrowserAvailable);
            Assert.False(string.IsNullOrWhiteSpace(capability.BrowserLabel));
        });
    }

    [Fact]
    public void ProvenanceCleanupRemovesSampleDeclarationAndPreservesCompressedPixels() {
        byte[] input = System.IO.File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "provenance-demo.png"));
        byte[] original = input.ToArray();
        var before = OfficeProvenanceBufferWorkflow.Inspect(input, "sample.png");
        Assert.Equal(OfficeProvenanceCarrierKind.IptcDigitalSourceType, Assert.Single(before.Evidence).Carrier);
        var result = OfficeProvenanceBufferWorkflow.Remove(input, "sample.png", new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = false, RemoveExternalC2paReferences = false, RemoveAiSourceMetadata = true
        });
        Assert.True(result.WasChanged);
        Assert.Empty(OfficeProvenanceBufferWorkflow.Inspect(result.ToArray(), "copy.png").Evidence);
        Assert.Equal(original, input);
        Assert.Equal(ImageData(input), ImageData(result.ToArray()));
        var unchanged = OfficeProvenanceBufferWorkflow.Remove(input, "sample.png", new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = false, RemoveExternalC2paReferences = false, RemoveAiSourceMetadata = false
        });
        Assert.False(unchanged.WasChanged); Assert.Equal(input, unchanged.ToArray());
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceBufferWorkflow.Inspect(input, "renamed.jpg"));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceBufferWorkflow.Inspect(input, "large.png", new OfficeProvenanceOptions { MaxAssetBytes = 1, MaxManifestBytes = 1 }));
    }

    private static byte[] ImageData(byte[] png) {
        using var output = new MemoryStream();
        for (int offset = 8; offset < png.Length;) {
            int length = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(offset));
            if (System.Text.Encoding.ASCII.GetString(png, offset + 4, 4) == "IDAT") output.Write(png, offset + 8, length);
            offset += length + 12;
        }
        return output.ToArray();
    }
}
