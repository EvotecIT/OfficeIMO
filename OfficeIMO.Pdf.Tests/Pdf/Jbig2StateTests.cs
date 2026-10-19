using System.Text.Json;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class Jbig2StateTests {
    private static JsonDocument OpenCases() => JsonDocument.Parse(File.ReadAllText(Path.Combine(
        AppContext.BaseDirectory, "Pdf", "Fixtures", "Rendering", "jbig2-state.json")));

    [Fact]
    public void PageOperatorsAndRetainedContextsMatchIndependentDecodedPixels() {
        using JsonDocument fixture = OpenCases();
        foreach (JsonElement item in fixture.RootElement.GetProperty("cases").EnumerateArray()) {
            byte[] encoded = Convert.FromBase64String(item.GetProperty("encoded").GetString()!);
            byte[] expected = Convert.FromBase64String(item.GetProperty("decoded").GetString()!);
            byte[] actual = Jbig2Decoder.Decode(encoded, null, item.GetProperty("width").GetInt32(), item.GetProperty("height").GetInt32(), 1 << 20, default);
            Assert.True(expected.SequenceEqual(actual), item.GetProperty("name").GetString());
            Assert.Null(Jbig2Decoder.FindUnsupportedFeature(encoded));
        }
    }

    [Fact]
    public void PageDefaultWinsWhenDirectRegionOverridesAreDisabled() {
        using JsonDocument fixture = OpenCases();
        JsonElement item = fixture.RootElement.GetProperty("cases").EnumerateArray()
            .Single(value => value.GetProperty("name").GetString() == "generic-op1-override0");
        List<byte[]> segments = Split(Convert.FromBase64String(item.GetProperty("encoded").GetString()!));
        // Deliberately conflicting direct-region headers cannot bypass the page's AND operator.
        foreach (byte[] segment in segments.Skip(1)) segment[11 + 16] = 0;
        byte[] actual = Jbig2Decoder.Decode(segments.SelectMany(value => value).ToArray(), null, 8, 8, 1 << 20, default);
        Assert.Equal(Convert.FromBase64String(item.GetProperty("decoded").GetString()!), actual);
    }

    [Fact]
    public void RetainedDictionaryContextsCrossTheGlobalsBoundary() {
        using JsonDocument fixture = OpenCases();
        JsonElement item = fixture.RootElement.GetProperty("cases").EnumerateArray()
            .Single(value => value.GetProperty("name").GetString() == "retained-generic");
        List<byte[]> segments = Split(Convert.FromBase64String(item.GetProperty("encoded").GetString()!));
        byte[] globals = segments[1];
        globals[6] = 0; // Dictionary page association zero.
        segments.RemoveAt(1);
        byte[] actual = Jbig2Decoder.Decode(segments.SelectMany(value => value).ToArray(), globals, 24, 8, 1 << 20, default);
        Assert.Equal(Convert.FromBase64String(item.GetProperty("decoded").GetString()!), actual);
    }

    [Theory]
    [InlineData("not-retained")]
    [InlineData("different-adaptive-pixel")]
    [InlineData("last-dictionary-not-retained")]
    public void InvalidRetainedContextReferencesAreRejected(string mutation) {
        using JsonDocument fixture = OpenCases();
        JsonElement item = fixture.RootElement.GetProperty("cases").EnumerateArray()
            .Single(value => value.GetProperty("name").GetString() == "retained-generic");
        List<byte[]> segments = Split(Convert.FromBase64String(item.GetProperty("encoded").GetString()!));
        switch (mutation) {
            case "not-retained": segments[1][11] &= 0xFD; break; // Clear SDCTXRETAINED.
            case "different-adaptive-pixel": segments[2][12 + 2] ^= 1; break;
            case "last-dictionary-not-retained":
                // Child one did not retain its contexts. It must not fall back to the retained parent.
                byte[] child = segments[3];
                child[5] = 64; // Two referred dictionaries.
                segments[3] = child.Take(7).Concat(new byte[] { 3 }).Concat(child.Skip(7)).ToArray();
                break;
        }
        Assert.Throws<InvalidDataException>(() => Jbig2Decoder.Decode(
            segments.SelectMany(value => value).ToArray(), null, 24, 8, 1 << 20, default));
    }

    // Fixtures use short page associations and at most two one-byte segment references.
    private static List<byte[]> Split(byte[] encoded) {
        var result = new List<byte[]>();
        for (int position = 0; position < encoded.Length;) {
            int header = 11 + (encoded[position + 5] >> 5);
            int lengthOffset = position + header - 4;
            int length = (encoded[lengthOffset] << 24) | (encoded[lengthOffset + 1] << 16) |
                (encoded[lengthOffset + 2] << 8) | encoded[lengthOffset + 3];
            result.Add(encoded.Skip(position).Take(header + length).ToArray());
            position += header + length;
        }
        return result;
    }
}
