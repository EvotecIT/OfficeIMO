namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkNumbersArchiveBuilder {
    private readonly List<(ulong InfoId, ulong Lower, ulong Upper, int Rows, int Columns)> _calculationTables = new();

    private void AddCalculationGraph() {
        var owners = P();
        var tracker = P();
        uint ownerIndex = 1;
        foreach (var table in _calculationTables) {
            _cancellationToken.ThrowIfCancellationRequested();
            // Creation identities hold network-order UUID halves. TSCE's TSP UUID
            // halves and four-word owner map hold consecutive little-endian bytes.
            ulong lower = ReverseBytes(table.Upper), upper = ReverseBytes(table.Lower);
            owners.Message(1, P().UInt(1, ownerIndex).Message(2,
                P().UInt(2, (uint)lower).UInt(3, lower >> 32).UInt(4, (uint)upper).UInt(5, upper >> 32)));
            var range = P().UInt(1, 0).UInt(2, 0).UInt(3, (uint)table.Columns - 1).UInt(4, (uint)table.Rows - 1);
            var spanning = P().Message(2, range).Message(3, range);
            var volatileDependencies = P();
            foreach (int field in new[] { 1, 2, 3, 4, 5, 7 }) volatileDependencies.Message(field, P());
            var dependency = P().Message(1, P().UInt(1, lower).UInt(2, upper)).UInt(2, ownerIndex).UInt(3, 1)
                .Message(4, P()).Message(5, P()).Message(6, volatileDependencies)
                .Message(7, spanning).Message(8, spanning).Message(9, P().Message(1, P())).Message(10, P())
                .Reference(11, table.InfoId).Message(13, P()).Message(14, P()).Message(15, P()).Message(16, P());
            ulong dependencyId = Allocate();
            Add(dependencyId, 4008, dependency);
            tracker.Reference(6, dependencyId);
            ownerIndex++;
        }
        tracker.Message(3, owners).UInt(5, 0);
        Add(CalculationId, 4000, P().Message(2, tracker).Reference(3, NamesId)
            .UInt(4, 1).UInt(6, 0).Message(7, P()).Message(13, P()));
    }

    private static ulong ReverseBytes(ulong value) {
        ulong reversed = 0;
        for (int index = 0; index < 8; index++) { reversed = reversed << 8 | (byte)value; value >>= 8; }
        return reversed;
    }
}
