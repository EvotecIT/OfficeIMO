using OfficeIMO.Provenance;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserDocumentWorkflowTests {
    [Fact]
    public void SessionChainsResultsAndRestoresOriginalSelection() {
        var session = new BrowserDocumentSession();
        var first = File("first.pdf", [1, 2]); var second = File("second.pdf", [3]);
        session.Open([first, second]);
        session.SelectCurrent([second, first]);
        int revision = session.Revision;
        session.SetResult([4, 5], "merged.pdf", revision);
        session.UseResult();
        Assert.Equal("merged.pdf", Assert.Single(session.Current).Name);
        session.SetResult([9], "stale.pdf", revision);
        Assert.Null(session.LatestResult);
        session.RestoreOriginals();
        Assert.Equal(new[] { "first.pdf", "second.pdf" }, session.Current.Select(file => file.Name));
        session.SetResult([0], "pages.zip", session.Revision);
        Assert.Null(session.LatestResult);
        session.Clear();
        Assert.Empty(session.Current); Assert.Empty(session.Originals);
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
    private static SelectedDocument File(string name, byte[] bytes) => new(name, Path.GetExtension(name), "PDF", bytes.Length, bytes);
}
