using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsInterleavingTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void InterleavedMetadataPagesAndResourcesLoadAndSaveAsLogicalParts(XpsFormat format) {
        var doc = XpsDocument.Create(format); doc.AddPage(100, 80).AddPath("M0,0H20V20Z");
        doc.AddResource("Metadata/data.bin", new byte[] { 1, 2, 3, 4, 5 }, "application/octet-stream");
        byte[] interleaved = Interleave(doc.Save());
        var loaded = XpsDocument.Load(interleaved);
        Assert.Equal(doc.Pages[0].ToSvg().Svg, loaded.Pages[0].ToSvg().Svg);
        Assert.Equal(new byte[] { 1, 2, 3, 4, 5 }, loaded.GetPartBytes("Metadata/data.bin"));
        loaded.AddPage(120, 60); Assert.Equal(2, XpsDocument.Load(loaded.Save()).Pages.Count);
        Assert.DoesNotContain(loaded.PartNames, p => p.EndsWith(".piece", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("gap")]
    [InlineData("missing-last")]
    [InlineData("early-last")]
    [InlineData("duplicate-index")]
    [InlineData("mixed")]
    [InlineData("invalid-index")]
    public void RejectsAmbiguousOrIncompletePieceSets(string mode) {
        var doc = XpsDocument.Create(); doc.AddPage();
        var entries = new Dictionary<string, byte[]> {
            ["payload.bin/[0].piece"] = new byte[] { 1 },
            ["payload.bin/[1].last.piece"] = new byte[] { 2 }
        };
        switch (mode) {
            case "gap": entries.Remove("payload.bin/[0].piece"); break;
            case "missing-last": entries["payload.bin/[1].piece"] = entries["payload.bin/[1].last.piece"]; entries.Remove("payload.bin/[1].last.piece"); break;
            case "early-last": entries["payload.bin/[0].last.piece"] = entries["payload.bin/[0].piece"]; entries.Remove("payload.bin/[0].piece"); break;
            case "duplicate-index": entries["payload.bin/[00].piece"] = new byte[] { 3 }; break;
            case "mixed": entries["payload.bin"] = new byte[] { 3 }; break;
            case "invalid-index": entries["payload.bin/[-1].piece"] = new byte[] { 3 }; break;
        }
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(WithEntries(doc.Save(), entries)));
    }

    [Fact]
    public void AggregatePartLimitAppliesBeforePieceMaterializationAndPhysicalEntriesAreBounded() {
        var doc = XpsDocument.Create(); doc.AddPage(); doc.AddResource("payload.bin", new byte[4000], "application/octet-stream");
        byte[] interleaved = Interleave(doc.Save());
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(interleaved, new XpsReadOptions { MaximumPartBytes = 3000 }));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(interleaved, new XpsReadOptions { MaximumParts = 5 }));
    }
    internal static byte[] Interleave(byte[] input) {
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var pieces = new Dictionary<string, byte[]>();
        foreach (var entry in zip.Entries) {
            using var source = entry.Open(); using var data = new MemoryStream(); source.CopyTo(data); byte[] bytes = data.ToArray(); int half = bytes.Length / 2;
            pieces[entry.FullName + "/[0].piece"] = bytes.Take(half).ToArray();
            pieces[entry.FullName + "/[1].last.piece"] = bytes.Skip(half).ToArray();
        }
        // Reverse physical ordering proves logical piece numbers, not ZIP encounter order, control assembly.
        return WithEntries(Array.Empty<byte>(), pieces.Reverse().ToDictionary(p => p.Key, p => p.Value));
    }
    private static byte[] WithEntries(byte[] input, Dictionary<string, byte[]> entries) {
        using var output = new MemoryStream(); output.Write(input, 0, input.Length);
        using (var zip = new ZipArchive(output, input.Length == 0 ? ZipArchiveMode.Create : ZipArchiveMode.Update, true))
            foreach (var item in entries) { using var stream = zip.CreateEntry(item.Key).Open(); stream.Write(item.Value, 0, item.Value.Length); }
        return output.ToArray();
    }
}
