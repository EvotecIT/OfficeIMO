using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Read;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void XlsbTabularReader_DisposesItsRecordReaderWhenConstructionFails() {
        using var worksheetPart = new TrackingReadStream(new byte[] { 0x80 });

        Assert.Throws<EndOfStreamException>(() =>
            new XlsbTabularDataReader(
                worksheetPart,
                Array.Empty<string>(),
                Array.Empty<bool>(),
                uses1904DateSystem: false,
                hasHeaderRow: true,
                new ExcelReadOptions(),
                new XlsbImportOptions(),
                new XlsbRecordReadBudget(100),
                CancellationToken.None));

        Assert.True(worksheetPart.WasDisposed);
    }

    private sealed class TrackingReadStream : MemoryStream {
        internal TrackingReadStream(byte[] bytes) : base(bytes, writable: false) {
        }

        internal bool WasDisposed { get; private set; }

        protected override void Dispose(bool disposing) {
            WasDisposed = true;
            base.Dispose(disposing);
        }
    }
}
