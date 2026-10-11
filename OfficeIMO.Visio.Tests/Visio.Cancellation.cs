using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioCancellationTests {
    [Fact]
    public void CancellationDuringLegacyNamespaceNormalizationAbortsImport() {
        XDocument source = XDocument.Parse("<VisioDocument xmlns='urn:schemas-microsoft-com:office:visio'><Pages><Page ID='0'><Shapes/></Page></Pages></VisioDocument>");
        using var cancellation = new CancellationTokenSource();
        source.Changed += (_, _) => cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            VisioLegacyXmlCodec.ToPackage(source, VisioPackageType.Drawing,
                new VisioXmlConversionReport(), cancellation.Token));

        Assert.True(cancellation.IsCancellationRequested);
    }

    [Fact]
    public async Task CancellationAfterInputReadDoesNotPublishLoadedDocument() {
        VisioDocument source = VisioDocument.Create();
        source.AddPage("Page", 10, 7);
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAtEndOfInputStream(source.ToBytes(), cancellation);
        stream.Position = 3;

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            VisioDocument.LoadAsync(stream, cancellationToken: cancellation.Token));

        Assert.True(cancellation.IsCancellationRequested);
        Assert.True(stream.CanRead);
        Assert.Equal(3, stream.Position);
    }

    private sealed class CancelAtEndOfInputStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;

        internal CancelAtEndOfInputStream(byte[] bytes, CancellationTokenSource cancellation)
            : base(bytes, writable: false) {
            _cancellation = cancellation;
        }

        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) {
            int read = await base.ReadAsync(buffer, offset, count, cancellationToken);
            if (read == 0) _cancellation.Cancel();
            return read;
        }
    }
}
