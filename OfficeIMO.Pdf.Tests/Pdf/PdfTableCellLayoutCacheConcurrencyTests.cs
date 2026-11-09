using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTableCellLayoutCacheConcurrencyTests {
    [Fact]
    public async Task ConcurrentWritesOfOneMergedCellTableProduceTheSameOutput() {
        var rows = Enumerable.Range(0, 40).Select(row => row % 4 == 0
            ? new[] { PdfTableCell.Merge("Merged " + row, columnSpan: 2, rowSpan: 3), PdfTableCell.TextCell("C" + row) }
            : new[] { PdfTableCell.TextCell("A" + row), PdfTableCell.TextCell("B" + row), PdfTableCell.TextCell("C" + row) });
        PdfDocument document = PdfDocument.Create().Table(rows.ToArray());

        byte[][] outputs = await Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(() => document.ToBytes())));

        Assert.All(outputs, output => Assert.Equal(document.ToBytes(), output));
    }
}
