using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class FontNameMetadataTests {
    [Fact]
    public void FamilyIdentityAndCancellationApplyToTheAcceptedSnapshot() {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        string family = OfficeOpenTypeReader.TryCreate(data)!.ReadFamilyName()!;
        string key = OfficeFontNameMetadata.FamilyKey(family);
        Assert.True(OfficeFontNameMetadata.MatchesFamily(data, key, CancellationToken.None));
        Assert.False(OfficeFontNameMetadata.MatchesFamily(data, "unrelatedfamily", CancellationToken.None));
        using var canceled = new CancellationTokenSource();
        using var stream = new CancelAfterReadStream(data, canceled);
        Assert.Throws<OperationCanceledException>(() => OfficeFontNameMetadata.ReadFamilyKeys(stream, canceled.Token));
        Assert.Equal(1, stream.Reads);
        Assert.Throws<OperationCanceledException>(() => OfficeCloudFontFiles.EnumerateFamilyFiles(
            Path.GetTempPath(), key, canceled.Token).ToArray());
        Assert.Throws<OperationCanceledException>(() => new OfficeFontFaceCollection().TryAddInstalledFamily(
            family, OfficeFontFaceDescriptor.Regular, int.MaxValue, canceled.Token, out _, out _));
    }

    private sealed class CancelAfterReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancelAfterReadStream(byte[] data, CancellationTokenSource cancellation) : base(data, writable: false) {
            _cancellation = cancellation;
        }
        internal int Reads { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) {
            int result = base.Read(buffer, offset, count);
            Reads++;
            _cancellation.Cancel();
            return result;
        }
    }
}
