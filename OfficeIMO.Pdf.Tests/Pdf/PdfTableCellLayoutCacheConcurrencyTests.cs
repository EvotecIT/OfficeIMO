using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTableCellLayoutCacheConcurrencyTests {
    [Fact]
    public async Task ConcurrentWritesOfOneMergedCellTableProduceTheSameOutput() {
        // Merges start at row 1 so none crosses the default one-row header.
        var rows = Enumerable.Range(0, 40).Select(row => row % 4 == 1
            ? new[] { PdfTableCell.Merge("Merged " + row, columnSpan: 2, rowSpan: 3), PdfTableCell.TextCell("C" + row) }
            : new[] { PdfTableCell.TextCell("A" + row), PdfTableCell.TextCell("B" + row), PdfTableCell.TextCell("C" + row) }).ToArray();

        // A fresh document per round puts every round's writers on a cold layout cache at once.
        for (int round = 0; round < 20; round++) {
            PdfDocument document = PdfDocument.Create().Table(rows);
            byte[][] outputs = await Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(() => document.ToBytes())));
            Assert.All(outputs, output => Assert.Equal(outputs[0], output));
        }
    }
}
