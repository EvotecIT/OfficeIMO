namespace OfficeIMO.OneNote.Tests;

public sealed class BinaryPayloadTests {
    [Fact]
    public void FromBytesOwnsDefensiveCopies() {
        byte[] source = { 1, 2, 3 };
        OneNoteBinaryPayload payload = OneNoteBinaryPayload.FromBytes(source);
        source[0] = 9;

        byte[] first = payload.ToArray(3);
        first[1] = 9;
        byte[] second = payload.ToArray(3);

        Assert.Equal(new byte[] { 1, 2, 3 }, second);
    }

    [Fact]
    public void LazyPayloadOpensIndependentStreams() {
        int calls = 0;
        OneNoteBinaryPayload payload = OneNoteBinaryPayload.FromStreamFactory(() => {
            calls++;
            return new MemoryStream(new byte[] { 4, 5, 6 }, false);
        }, 3);

        Assert.Equal(new byte[] { 4, 5, 6 }, payload.ToArray(3));
        Assert.Equal(new byte[] { 4, 5, 6 }, payload.ToArray(3));
        Assert.Equal(2, calls);
    }

    [Fact]
    public void MaterializationLimitIsEnforcedBeforeOpeningKnownPayload() {
        bool opened = false;
        OneNoteBinaryPayload payload = OneNoteBinaryPayload.FromStreamFactory(() => {
            opened = true;
            return new MemoryStream(new byte[4], false);
        }, 4);

        Assert.Throws<IOException>(() => payload.ToArray(3));
        Assert.False(opened);
    }
}
