using OfficeIMO.Email.Store;
using Xunit;

namespace OfficeIMO.Email.Tests.Store.Emlx;

public class EmlxBinaryPlistReaderTests {
    [Fact]
    public void AliasedDataReferencesShareTheCumulativeDecodedByteBudget() {
        // Three distinct references share the same four-byte source object.
        byte[] plist = new byte[62];
        "bplist00"u8.CopyTo(plist);
        byte[] objects = { 0xD1, 1, 2, 0x53, (byte)'k', (byte)'e', (byte)'y',
            0xA3, 3, 4, 5, 0x44, 1, 2, 3, 4 };
        objects.CopyTo(plist, 8);
        byte[] offsets = { 8, 11, 15, 19, 19, 19 };
        offsets.CopyTo(plist, 24);
        int trailer = plist.Length - 32;
        plist[trailer + 6] = 1;
        plist[trailer + 7] = 1;
        plist[trailer + 15] = 6;
        plist[trailer + 31] = 24;

        var bounded = new EmlxBinaryPlistReader(plist, 0, EmailStoreReaderOptions.Default,
            CancellationToken.None, maxDecodedBytes: 28);
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem),
            Assert.Throws<EmailStoreLimitExceededException>(() => bounded.Read()).LimitName);

        var allowed = new EmlxBinaryPlistReader(plist, 0, EmailStoreReaderOptions.Default,
            CancellationToken.None, maxDecodedBytes: 42);
        var values = allowed.Read();
        object?[] data = Assert.IsType<object?[]>(values["key"]);
        Assert.Equal(3, data.Length);
        Assert.All(data, value => Assert.Equal(new byte[] { 1, 2, 3, 4 }, Assert.IsType<byte[]>(value)));
        Assert.Equal(18 + 3 * IntPtr.Size, allowed.DecodedBytes);
    }

    [Fact]
    public void DateRealsKeepSubMillisecondTicks() {
        // { "date": 2^-10 seconds after the 2001 epoch }, which is 9765.625 ticks.
        byte[] plist = new byte[60];
        "bplist00"u8.CopyTo(plist);
        byte[] objects = { 0xD1, 1, 2, 0x54, (byte)'d', (byte)'a', (byte)'t', (byte)'e', 0x33 };
        objects.CopyTo(plist, 8);
        byte[] seconds = BitConverter.GetBytes(1D / 1024D);
        if (BitConverter.IsLittleEndian) Array.Reverse(seconds);
        seconds.CopyTo(plist, 17);
        byte[] offsets = { 8, 11, 16 };
        offsets.CopyTo(plist, 25);
        int trailer = plist.Length - 32;
        plist[trailer + 6] = 1;
        plist[trailer + 7] = 1;
        plist[trailer + 15] = 3;
        plist[trailer + 31] = 25;

        var values = new EmlxBinaryPlistReader(plist, 0, EmailStoreReaderOptions.Default, CancellationToken.None, maxDecodedBytes: 1024).Read();

        Assert.Equal(new DateTimeOffset(2001, 1, 1, 0, 0, 0, TimeSpan.Zero).AddTicks(9765), values["date"]);
    }
}
