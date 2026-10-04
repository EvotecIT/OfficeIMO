using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfNameTreeByteOrderingTests {
    [Fact]
    public void GeneratedAndEditedAttachmentsSortEncodedKeysAndKeepPayloads() {
        string[] names = { "ÿ", "’", "a1", "a", "space\u00A0key" };
        byte[] bytes = PdfDocument.Create().Paragraph(p => p.Text("Attachments"))
            .AttachFile(names[0], new byte[] { 1 })
            .AttachFile(names[1], new byte[] { 2 })
            .AttachFile(names[2], new byte[] { 3 })
            .AttachFile(names[3], new byte[] { 4 })
            .AttachFile(names[4], new byte[] { 5 }).ToBytes();
        AssertOrderedKeys(bytes, "EmbeddedFiles");
        bytes = PdfAttachmentEditor.Add(bytes, new PdfEmbeddedFile("’extra", new byte[] { 6 })).ToBytes();
        AssertOrderedKeys(bytes, "EmbeddedFiles");
        Assert.Equal(6, PdfAttachmentExtractor.ExtractAttachments(bytes).Count);
        Assert.Equal(new byte[] { 6 }, Assert.Single(PdfAttachmentExtractor.ExtractAttachments(bytes), a => a.FileName == "’extra").Bytes);
    }

    [Fact]
    public void MergedDestinationsSortEncodedKeysAndKeepPageTargets() {
        byte[] first = PdfDocument.Create().Canvas(c => c.NamedDestination("ÿ", 10, 10)).ToBytes();
        byte[] second = PdfDocument.Create().Canvas(c => c.NamedDestination("’", 20, 20).NamedDestination("space\u00A0key", 30, 30)).ToBytes();
        byte[] merged = PdfMerger.MergeResult(new PdfMergeOptions {
            Policy = new PdfMergePolicy { NamedDestinations = PdfMergeStructureMode.Combine }
        }, first, second).ToBytes();
        AssertOrderedKeys(merged, "Dests");
        var info = PdfInspector.Inspect(merged);
        Assert.Equal(1, Assert.Single(info.NamedDestinations, d => d.Name == "ÿ").PageNumber);
        Assert.Equal(2, Assert.Single(info.NamedDestinations, d => d.Name == "’").PageNumber);
        Assert.Contains("space\u00A0key", info.NamedDestinationNames);
    }

    private static void AssertOrderedKeys(byte[] bytes, string treeName) {
        var map = PdfSyntax.ParseObjects(bytes).Map;
        PdfDictionary catalog = Assert.Single(map.Values.Select(o => o.Value).OfType<PdfDictionary>(),
            d => d.Items.TryGetValue("Type", out var type) && type is PdfName { Name: "Catalog" });
        PdfObject Resolve(PdfObject value) => value is PdfReference r ? map[r.ObjectNumber].Value : value;
        var names = Assert.IsType<PdfDictionary>(Resolve(catalog.Items["Names"]));
        var tree = Assert.IsType<PdfDictionary>(Resolve(names.Items[treeName]));
        var entries = Assert.IsType<PdfArray>(tree.Items["Names"]);
        for (int index = 2; index < entries.Items.Count; index += 2) {
            byte[] previous = Assert.IsType<PdfStringObj>(entries.Items[index - 2]).RawBytes;
            byte[] current = Assert.IsType<PdfStringObj>(entries.Items[index]).RawBytes;
            // Independent lexicographic check uses fixed-width hex, including shorter prefixes first.
            Assert.True(StringComparer.Ordinal.Compare(BitConverter.ToString(previous), BitConverter.ToString(current)) < 0);
        }
    }
}
