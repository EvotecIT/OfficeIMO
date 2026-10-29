using System.Collections.Generic;
using OfficeIMO.Reader;
using OfficeIMO.Reader.DjVu;

namespace OfficeIMO.DjVu.Tests;

public sealed class ReaderTextProjectionLimitsTests {
    [Theory]
    [InlineData("direct")]
    [InlineData("document")]
    [InlineData("chunks")]
    public void OverlappingWordZonesRespectTheAggregateTextLimit(string entryPoint) {
        byte[] source = PageWithText();
        var options = new ReaderDjVuOptions {
            ReadOptions = new DjVuReadOptions { MaxTextCharacters = 9 }
        };
        var document = DjVuDocument.Load(source);
        Assert.Equal(DjVuTextStatus.Present, document.Pages[0].GetText().Status);
        var accepted = document.ToReadResult(options);
        Assert.Equal(new[] { "abc", "abc", "abc" }, accepted.Blocks.Select(block => block.Text));

        options.ReadOptions.MaxTextCharacters = 8;
        var reader = new OfficeDocumentReaderBuilder().AddDjVuHandler(options).Build();
        var error = Assert.Throws<ReaderResourceLimitException>(() => {
            if (entryPoint == "direct") document.ToReadResult(options);
            else if (entryPoint == "document") reader.ReadDocument(source, "words.djvu");
            else reader.Read(source, "words.djvu").ToArray();
        });
        Assert.Equal(nameof(DjVuReadOptions.MaxTextCharacters), error.LimitName);
        Assert.Equal(8, error.Maximum);
    }

    [Theory]
    [InlineData(true, 18)]
    [InlineData(false, 8)]
    public void ProjectionTextLimitIsSharedAcrossSelectedPages(bool wordZones, int acceptedCharacters) {
        byte[] index = File.ReadAllBytes(ReaderAdapterTests.Fixture(Path.Combine("Shared", "Indirect", "index.djvu")));
        var components = SharedComponentTests.Components();
        for (int page = 1; page <= 2; page++) {
            string id = "shared-" + page + ".djvu";
            components[id] = SharedComponentTests.AppendChunk(components[id], "TXTa", TextPayload(wordZones));
        }
        var document = DjVuDocument.Load(index, new DjVuReadOptions {
            ComponentResolver = (id, _) => components[id]
        });
        var options = new ReaderDjVuOptions {
            ReadOptions = new DjVuReadOptions { MaxTextCharacters = acceptedCharacters }
        };
        Assert.Equal(acceptedCharacters, document.ToReadResult(options).Blocks.Sum(block => block.Text!.Length));
        options.ReadOptions.MaxTextCharacters--;
        var error = Assert.Throws<ReaderResourceLimitException>(() => document.ToReadResult(options));
        Assert.Equal(nameof(DjVuReadOptions.MaxTextCharacters), error.LimitName);
        Assert.Equal(acceptedCharacters - 1, error.Maximum);
    }

    private static byte[] PageWithText() => SharedComponentTests.AppendChunk(
        File.ReadAllBytes(ReaderAdapterTests.Fixture("fax.djvu")), "TXTa", TextPayload(true));

    private static byte[] TextPayload(bool wordZones) {
        var bytes = new List<byte> { 0, 0, 4, (byte)'a', (byte)'b', (byte)'c', (byte)'d' };
        if (wordZones) {
            bytes.Add(1); // Text-zone version.
            Zone(1, 0, 4, 3);
            Zone(6, 0, 3, 0);
            Zone(6, -3, 3, 0);
            Zone(6, -3, 3, 0);
        }
        return bytes.ToArray();

        void Zone(byte kind, int relativeText, int length, int children) {
            bytes.Add(kind);
            Biased16(0); Biased16(0); Biased16(1); Biased16(1); Biased16(relativeText);
            U24(length); U24(children);
        }
        void Biased16(int value) {
            int encoded = value + 32768;
            bytes.Add((byte)(encoded >> 8)); bytes.Add((byte)encoded);
        }
        void U24(int value) {
            bytes.Add((byte)(value >> 16)); bytes.Add((byte)(value >> 8)); bytes.Add((byte)value);
        }
    }
}
