using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class EarlyBiffBufferingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LoadCancellationAfterTheBiffProbeStopsBeforeTheBody(bool seekable, bool securityEnabled) {
        byte[] worksheet = CreateProbeWorksheet();
        var options = new ExcelLoadOptions {
            PackageSecurity = securityEnabled
                ? new OfficePackageSecurityOptions { MaxPackageBytes = 1024 * 1024 }
                : null
        };
        using (var valid = new MemoryStream(worksheet, writable: false))
        using (ExcelDocument document = ExcelDocument.Load(valid, options, CancellationToken.None)) {
            Assert.Equal(42.25d, document.Sheets[0].Cell(1, 1).GetValue<double>());
        }

        using var cancellation = new CancellationTokenSource();
        using var source = new ProbeReadStream(worksheet, cancellation, seekable);
        if (seekable) source.Position = 3;

        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(
            () => ExcelDocument.Load(source, options, cancellation.Token));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(0, source.LengthReadsAfterCancellation);
        Assert.Equal(0, source.ReadAttemptsAfterCancellation);
        Assert.Equal(10, source.BytesRead);
        Assert.Equal(1, source.ReadCalls);
        Assert.True(source.CanRead);
        if (seekable) Assert.Equal(3, source.Position);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task LoadAsyncCancellationAfterTheBiffProbeStopsBeforeFurtherAdmission(bool seekable, bool securityEnabled) {
        byte[] worksheet = CreateProbeWorksheet();
        var options = new ExcelLoadOptions {
            PackageSecurity = securityEnabled
                ? new OfficePackageSecurityOptions { MaxPackageBytes = 1024 * 1024 }
                : null
        };
        using (var valid = new MemoryStream(worksheet, writable: false))
        using (ExcelDocument document = await ExcelDocument.LoadAsync(valid, options, CancellationToken.None)) {
            Assert.Equal(42.25d, document.Sheets[0].Cell(1, 1).GetValue<double>());
        }

        using var cancellation = new CancellationTokenSource();
        using var source = new ProbeReadStream(worksheet, cancellation, seekable);
        if (seekable) source.Position = 3;

        OperationCanceledException error = await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => ExcelDocument.LoadAsync(source, options, cancellation.Token));

        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(0, source.LengthReadsAfterCancellation);
        Assert.Equal(0, source.ReadAttemptsAfterCancellation);
        Assert.Equal(10, source.BytesRead);
        Assert.Equal(1, source.ReadCalls);
        Assert.True(source.CanRead);
        if (seekable) Assert.Equal(3, source.Position);
    }

    [Theory]
    [InlineData(false, false, 24, 64)]
    [InlineData(false, true, 24, 64)]
    [InlineData(true, false, 24, 64)]
    [InlineData(true, true, 24, 64)]
    [InlineData(false, false, 64, 24)]
    [InlineData(false, true, 64, 24)]
    [InlineData(true, false, 64, 24)]
    [InlineData(true, true, 64, 24)]
    [InlineData(false, false, 24, 24)]
    [InlineData(false, true, 24, 24)]
    [InlineData(true, false, 24, 24)]
    [InlineData(true, true, 24, 24)]
    public async Task LoadReportsTheBindingFormatOrPackageLimit(bool asynchronous, bool seekable, int legacyLimit, int packageLimit) {
        using var cancellation = new CancellationTokenSource();
        using var source = new ProbeReadStream(CreateProbeWorksheet(), cancellation, seekable, cancelOnFirstRead: false);
        if (seekable) source.Position = 3;
        var options = new ExcelLoadOptions {
            MaxLegacyInputBytes = legacyLimit,
            PackageSecurity = new OfficePackageSecurityOptions { MaxPackageBytes = packageLimit }
        };
        Exception? failure = await Record.ExceptionAsync(async () => {
            using ExcelDocument document = asynchronous
                ? await ExcelDocument.LoadAsync(source, options)
                : ExcelDocument.Load(source, options);
        });
        if (packageLimit <= legacyLimit) {
            OfficePackageSecurityException package = Assert.IsType<OfficePackageSecurityException>(failure);
            Assert.Equal(OfficePackageSecurityRule.PackageSize, package.Rule);
            Assert.Equal(24D, package.Limit);
        } else {
            InvalidDataException format = Assert.IsType<InvalidDataException>(failure);
            Assert.Contains("24 bytes", format.Message);
        }
        Assert.Equal(seekable ? packageLimit <= legacyLimit ? 0 : 10 : 25, source.BytesRead);
        Assert.True(source.CanRead);
        if (seekable) Assert.Equal(3, source.Position);
    }

    // BIFF4 BOF, NUMBER(A1 = 42.25), EOF. The complete BOF occupies ten bytes.
    private static byte[] CreateProbeWorksheet() => new byte[] {
        0x09, 0x04, 0x06, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00,
        0x03, 0x02, 0x0E, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x20, 0x45, 0x40,
        0x0A, 0x00, 0x00, 0x00
    };

    private sealed class ProbeReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        private readonly bool _seekable;
        private readonly bool _cancelOnFirstRead;

        internal ProbeReadStream(byte[] bytes, CancellationTokenSource cancellation, bool seekable, bool cancelOnFirstRead = true)
            : base(bytes, writable: false) {
            _cancellation = cancellation;
            _seekable = seekable;
            _cancelOnFirstRead = cancelOnFirstRead;
        }

        public override bool CanSeek => _seekable && base.CanSeek;
        internal int BytesRead { get; private set; }
        internal int ReadCalls { get; private set; }
        internal int ReadAttemptsAfterCancellation { get; private set; }
        internal int LengthReadsAfterCancellation { get; private set; }

        public override long Length {
            get {
                if (_cancellation?.IsCancellationRequested == true) LengthReadsAfterCancellation++;
                return base.Length;
            }
        }

        public override int Read(byte[] buffer, int offset, int count) {
            if (_cancellation.IsCancellationRequested) ReadAttemptsAfterCancellation++;
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            if (++ReadCalls == 1 && _cancelOnFirstRead) _cancellation.Cancel();
            return read;
        }

        // A completed I/O may return bytes normally while cancellation becomes set.
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            Task.FromResult(Read(buffer, offset, count));
    }
}
