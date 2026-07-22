using OfficeIMO.Email;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstWriterTableTests {
    [Fact]
    public void Folder_table_row_matrix_uses_the_table_budget_by_default() {
        var options = new EmailStoreReaderOptions(
            maxDecodedPropertyBytesPerItem: 128,
            maxDecodedTableBytes: 4096);

        Assert.Equal(4096, PstTableContextReader.GetMaximumRowMatrixBytes(options, null));
        Assert.Equal(64, PstTableContextReader.GetMaximumRowMatrixBytes(options, 64));
    }

    [Fact]
    public void Multi_block_heap_and_row_index_round_trip_more_than_one_thousand_rows() {
        string path = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-pst-table-", Guid.NewGuid().ToString("N"), ".pst"));
        try {
            var rows = Enumerable.Range(0, 1_500).Select(index =>
                new PstWriterTableRow(checked((uint)(0x2004 + index * 0x20)), new[] {
                    new MapiProperty(0x0037, MapiPropertyType.Unicode,
                        string.Concat("Subject ", index.ToString(CultureInfo.InvariantCulture)))
                })).ToArray();
            using (var writer = new PstWriterFile(path)) {
                PstWriterContextResult table = PstTableContextWriter.Write(writer,
                    rows, 65001,
                    new[] { new MapiProperty(0x0037, MapiPropertyType.Unicode) },
                    null, "large-table");
                var nodes = new[] { new PstWriterNode(0x60E, 0, table.DataBid, table.SubnodeBid) };
                PstWriterTreeRoot nbt = writer.WriteNodeTree(nodes);
                PstWriterTreeRoot bbt = writer.WriteBlockTree();
                writer.FinalizeFile(nbt, bbt, nodes);
            }

            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            PstHeader header = PstHeader.Read(stream, EmailStoreFormat.Pst);
            var ndb = new PstNdbReader(stream, header, EmailStoreReaderOptions.Default,
                CancellationToken.None);
            ndb.LoadIndexes();
            PstNodeReference node = ndb.Nodes[0x60E];
            PstDataTree data = ndb.ReadDataTree(node.DataBid,
                64 * 1024 * 1024, CancellationToken.None);
            IReadOnlyDictionary<uint, PstSubnodeReference> subnodes =
                ndb.ReadSubnodes(node.SubnodeBid, CancellationToken.None);
            var heap = new PstHeap(data, subnodes, ndb,
                EmailStoreReaderOptions.Default, CancellationToken.None);
            IReadOnlyList<IReadOnlyList<MapiProperty>> decoded =
                new PstTableContextReader(heap, true,
                    EmailStoreReaderOptions.Default, CancellationToken.None).ReadRows();

            Assert.Equal(rows.Length, decoded.Count);
            Assert.Equal("Subject 1499", decoded[1499]
                .Single(item => item.PropertyId == 0x0037).Value);
        } finally {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
