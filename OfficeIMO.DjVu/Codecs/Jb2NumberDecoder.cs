namespace OfficeIMO.DjVu;

// Appendix 2's bounded sign/range/value decision tree. Inferred decisions still
// traverse the same tree so later, less constrained values reuse the correct contexts.
internal sealed class Jb2NumberDecoder {
    private sealed class Node {
        internal byte Context;
        internal Node? Left, Right;
    }
    private readonly ZpDecoder _arithmetic;
    private readonly DjVuReadBudget _budget;
    private readonly Node[] _roots = new Node[16];
    private int _nodes;
    private readonly Func<long> _otherBuffers;
    internal long RetainedBytes => _nodes * 40L;
    internal Jb2NumberDecoder(ZpDecoder arithmetic, DjVuReadBudget budget, Func<long> otherBuffers) { _arithmetic = arithmetic; _budget = budget; _otherBuffers = otherBuffers; Reset(); }
    internal void Reset() { for (int i = 0; i < _roots.Length; i++) _roots[i] = new Node(); _nodes = _roots.Length; }

    internal int Read(int context, int low, int high) {
        if (low > high) throw new InvalidDataException("Invalid JB2 integer range.");
        Node node = _roots[context];
        int positive = low >= 0 ? 1 : high < 0 ? 0 : _arithmetic.Bit(ref node.Context);
        node = Child(node, positive);
        if (positive == 0) { int oldLow = low; low = -high - 1; high = -oldLow - 1; }
        low = Math.Max(low, 0);
        int first = 0, width = 1;
        while (true) {
            int last = first + width - 1;
            int beyond = low > last ? 1 : high <= last ? 0 : _arithmetic.Bit(ref node.Context);
            node = Child(node, beyond);
            if (beyond == 0) break;
            first += width;
            width <<= 1;
            if (width > 262144) throw new InvalidDataException("JB2 integer exceeds the format range.");
        }
        while (width > 1) {
            width >>= 1;
            int middle = first + width;
            int upper = low >= middle ? 1 : high < middle ? 0 : _arithmetic.Bit(ref node.Context);
            node = Child(node, upper);
            if (upper != 0) first = middle;
        }
        if (first < low || first > high) throw new InvalidDataException("JB2 integer is outside its declared range.");
        return positive == 0 ? -first - 1 : first;
    }

    private Node Child(Node node, int bit) {
        if (bit == 0 && node.Left != null) return node.Left;
        if (bit != 0 && node.Right != null) return node.Right;
        _budget.WorkingBytes(++_nodes * 40L + _otherBuffers());
        var child = new Node();
        if (bit == 0) node.Left = child; else node.Right = child;
        return child;
    }
}
