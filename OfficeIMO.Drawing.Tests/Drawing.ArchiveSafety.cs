using OfficeIMO.Drawing.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingArchiveSafetyTests {
    [Fact]
    public void EntryCopyStopsAtDeclaredExpansionBound() {
        using var source = new CountingReadStream(new byte[1024]);

        InvalidDataException exception = Assert.Throws<
            InvalidDataException>(() => OfficeArchiveSafety.ReadEntryBytes(
                source, declaredLength: 16, maximumLength: 1024));

        Assert.Contains("declared expansion length", exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(17, source.BytesRead);
    }

    [Fact]
    public void EntryCopyRejectsTruncatedPayload() {
        using var source = new CountingReadStream(new byte[8]);

        InvalidDataException exception = Assert.Throws<
            InvalidDataException>(() => OfficeArchiveSafety.ReadEntryBytes(
                source, declaredLength: 16, maximumLength: 1024));

        Assert.Contains("shorter than its declared length", exception.Message,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(8, source.BytesRead);
    }

    private sealed class CountingReadStream : MemoryStream {
        public CountingReadStream(byte[] bytes) : base(bytes,
            writable: false) {
        }

        public int BytesRead { get; private set; }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }
    }
}
