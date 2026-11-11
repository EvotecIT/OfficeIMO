using System.Buffers.Binary;
using System.Text;
using System.IO.Packaging;
using OfficeIMO.Web.Converter.Engine;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class OriginAttributionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AiDeclarationNamesOnlyItsOwnCredentialGenerator(bool separateAiCredential) {
        var session = new ToolSession();
        session.Stage(0, CreateInput(separateAiCredential), separateAiCredential ? "mixed.docx" : "mixed.png");
        var result = ShareTool.Run(session, "inspect", new ToolOptions(new()));
        Assert.Contains("generative AI", result.Verdict.Detail);
        Assert.DoesNotContain("generative AI tool: Photoshop", result.Verdict.Detail);
        if (separateAiCredential) {
            Assert.Contains("generative AI tool: AI Studio", result.Verdict.Detail);
            Assert.Contains("embedded asset", result.Verdict.Detail);
            Assert.DoesNotContain("The file says it was made", result.Verdict.Detail);
            Assert.Contains(result.Items, item => (item.Group ?? "").Contains("word/media/ai.png"));
        }
        else Assert.DoesNotContain("generative AI tool:", result.Verdict.Detail);
    }

    [Fact]
    public void EmbeddedAssetPathRetainsDirectoriesNamedLikeFormatMarkers() {
        byte[] ai = RewriteCredential(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "unsigned-credential-12-actions.png")), ("Generator", "AI Studio"));
        var session = new ToolSession(); session.Stage(0, WithImages(("word/media/PNG/ai.png", ai)), "nested.docx");
        var result = ShareTool.Run(session, "inspect", new ToolOptions(new()));
        Assert.Contains("embedded asset (word/media/PNG/ai.png)", result.Verdict.Detail);
        Assert.Contains(result.Items, item => (item.Group ?? "").Contains("word/media/PNG/ai.png"));
    }

    [Theory]
    [InlineData("digitalCapture", false)]
    [InlineData("algorithmicMedia", false)]
    [InlineData("compositeCapture", false)]
    [InlineData("digitalCapture", true)]
    [InlineData("algorithmicMedia", true)]
    [InlineData("compositeCapture", true)]
    public void OrdinarySourceLabelsRemainInformationalWhenAiLabelsAreRemoved(string ordinary, bool includeAi) {
        var session = new ToolSession();
        session.Stage(0, SourceLabelInput(ordinary, includeAi), "sources.png");
        var inspection = ShareTool.Run(session, "inspect", new ToolOptions(new()));
        Assert.Equal(includeAi ? 1 : 0, inspection.Items.Count(item => item.Selectable));
        Assert.Contains(inspection.Items, item => !item.Selectable && item.Detail.Contains("camera", StringComparison.OrdinalIgnoreCase) ||
            !item.Selectable && item.Detail.Contains("software", StringComparison.OrdinalIgnoreCase) ||
            !item.Selectable && item.Detail.Contains("captured", StringComparison.OrdinalIgnoreCase));
        var removal = ShareTool.Run(session, "remove", new ToolOptions(new() { ["remove"] = "declarations" }));
        Assert.DoesNotContain(removal.Items, item => item.State == ToolState.Bad);
        Assert.Contains(removal.Items, item => item.State == ToolState.Kept && item.Title.Contains("source label", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(includeAi ? "1" : "0", removal.Facts.Single(fact => fact.Label == "Removed").Value);
        Assert.Equal("1", removal.Facts.Single(fact => fact.Label == "Still present").Value);
    }

    [Fact]
    public void RemovalPreservesWatermarkWarningBeyondDisplayedCredentials() {
        var session = new ToolSession();
        session.Stage(0, WatermarkInput(), "four-credentials.docx");
        var removal = ShareTool.Run(session, "remove", new ToolOptions(new() { ["remove"] = "provenance" }));
        Assert.Contains(removal.Items, item => item.Id == "watermark" && item.State == ToolState.Warning);
        Assert.Contains("watermark", removal.Verdict.Detail);
    }

    private static byte[] WatermarkInput() {
        byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "unsigned-credential-12-actions.png"));
        byte[] plain = RewriteCredential(original, ("c2pa.watermarked.bound", "c2pa.edited".PadRight("c2pa.watermarked.bound".Length)));
        return WithImages(Enumerable.Range(0, 4).Select(index => ($"word/media/credential-{index}.png", index == 3 ? original : plain)).ToArray());
    }

    /// <summary>The business-summary sample with each picture added as an image part related to the main document, as Word stores it.</summary>
    private static byte[] WithImages(params (string Path, byte[] Png)[] images) {
        using var output = new MemoryStream();
        output.Write(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "business-summary.docx")));
        using (Package package = Package.Open(output, FileMode.Open, FileAccess.ReadWrite)) {
            PackagePart main = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
            foreach ((string path, byte[] png) in images) {
                var uri = new Uri("/" + path, UriKind.Relative);
                using (Stream stream = package.CreatePart(uri, "image/png").GetStream()) stream.Write(png);
                main.CreateRelationship(uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
            }
        }
        return output.ToArray();
    }

    private static byte[] SourceLabelInput(string ordinary, bool includeAi) {
        byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "unsigned-credential-12-actions.png"));
        const string prefix = "http://cv.iptc.org/newscodes/digitalsourcetype/";
        string Description(string value) => "<rdf:Description xmlns:i='http://iptc.org/std/Iptc4xmpExt/2008-02-29/'>" +
            $"<i:DigitalSourceType>{prefix}{value}</i:DigitalSourceType></rdf:Description>";
        string xmp = "<x:xmpmeta xmlns:x='adobe:ns:meta/'><rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'>" +
            Description(ordinary) + (includeAi ? Description("trainedAlgorithmicMedia") : "") + "</rdf:RDF></x:xmpmeta>";
        byte[] content = Encoding.UTF8.GetBytes("XML:com.adobe.xmp\0\0\0\0\0" + xmp);
        byte[] chunk = new byte[content.Length + 12];
        BinaryPrimitives.WriteInt32BigEndian(chunk, content.Length);
        Encoding.ASCII.GetBytes("iTXt").CopyTo(chunk, 4); content.CopyTo(chunk, 8); SetCrc(chunk);
        using var output = new MemoryStream(); output.Write(original, 0, 8);
        foreach (byte[] imageChunk in Chunks(original)) {
            string type = Encoding.ASCII.GetString(imageChunk, 4, 4);
            if (type is "caBX" or "iTXt") continue;
            if (type == "IEND") output.Write(chunk);
            output.Write(imageChunk);
        }
        return output.ToArray();
    }

    private static void SetCrc(byte[] chunk) {
        uint crc = uint.MaxValue;
        for (int i = 4; i < chunk.Length - 4; i++) {
            crc ^= chunk[i];
            for (int bit = 0; bit < 8; bit++) crc = (crc >> 1) ^ ((crc & 1) == 0 ? 0 : 0xEDB88320u);
        }
        BinaryPrimitives.WriteUInt32BigEndian(chunk.AsSpan(chunk.Length - 4), ~crc);
    }

    private static byte[] CreateInput(bool separateAiCredential) {
        byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "unsigned-credential-12-actions.png"));
        byte[] ordinary = RewriteCredential(original, ("Generator", "Photoshop"),
            ("trainedAlgorithmicMedia", new string('x', "trainedAlgorithmicMedia".Length)));
        byte[] ai = separateAiCredential
            ? RewriteCredential(original, ("Generator", "AI Studio"))
            : File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "provenance-demo.png"));
        using var output = new MemoryStream();
        if (separateAiCredential) {
            return WithImages(("word/media/ordinary.png", ordinary), ("word/media/ai.png", ai));
        } else {
            output.Write(ordinary, 0, ordinary.Length - 12); // Before IEND; XMP may follow image data.
            foreach (byte[] chunk in Chunks(ai).Where(chunk => Encoding.ASCII.GetString(chunk, 4, 4) == "iTXt")) output.Write(chunk);
            output.Write(ordinary, ordinary.Length - 12, 12);
        }
        return output.ToArray();
    }

    private static IEnumerable<byte[]> Chunks(byte[] png) {
        for (int offset = 8; offset < png.Length;) {
            int size = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(offset)) + 12;
            yield return png.AsSpan(offset, size).ToArray();
            offset += size;
        }
    }

    private static byte[] RewriteCredential(byte[] png, params (string From, string To)[] replacements) {
        using var output = new MemoryStream();
        output.Write(png, 0, 8);
        foreach (byte[] chunk in Chunks(png)) {
            if (Encoding.ASCII.GetString(chunk, 4, 4) == "caBX") {
                foreach (var (from, to) in replacements) {
                    Assert.Equal(from.Length, to.Length); // Preserve CBOR string and JUMBF box lengths.
                    byte[] source = Encoding.UTF8.GetBytes(from), target = Encoding.UTF8.GetBytes(to);
                    for (int i = 8; i <= chunk.Length - 4 - source.Length; i++)
                        if (chunk.AsSpan(i, source.Length).SequenceEqual(source)) target.CopyTo(chunk, i);
                }
                SetCrc(chunk);
            }
            output.Write(chunk);
        }
        return output.ToArray();
    }
}
