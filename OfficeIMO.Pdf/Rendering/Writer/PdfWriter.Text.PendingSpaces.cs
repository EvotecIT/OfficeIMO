namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Keeps pending source styles in order without shifting or rescanning the remaining fragments.</summary>
    private sealed class PendingSpaceFragments : IReadOnlyList<PendingSpaceFragment> {
        private readonly List<PendingSpaceFragment> _items = new();
        private int _head;
        private double _advance;
        private double _correction;

        public int Count => _items.Count - _head;
        internal double TotalAdvance => _advance + _correction;

        public PendingSpaceFragment this[int index] {
            get => _items[_head + index];
            set {
                AddAdvance(-_items[_head + index].Advance);
                _items[_head + index] = value;
                AddAdvance(value.Advance);
            }
        }

        internal void Add(PendingSpaceFragment fragment) {
            AddAdvance(fragment.Advance);
            _items.Add(fragment);
        }

        internal void AddSpaceToLast() {
            PendingSpaceFragment fragment = _items[_items.Count - 1];
            fragment.AddSpace();
            AddAdvance(fragment.Style.LeadingAdvance);
        }

        internal void ConsumeFirstAdvance(double consumed) {
            _items[_head].Advance -= consumed;
            AddAdvance(-consumed);
        }

        internal void RemoveFirst() {
            AddAdvance(-_items[_head++].Advance);
            if (_head == _items.Count) Clear();
        }

        internal void Clear() {
            _items.Clear(); _head = 0; _advance = _correction = 0D;
        }

        private void AddAdvance(double amount) {
            // Retain small residual advances when positive and negative tracking cancel.
            double next = _advance + amount;
            _correction += Math.Abs(_advance) >= Math.Abs(amount)
                ? (_advance - next) + amount : (amount - next) + _advance;
            _advance = next;
            if (double.IsNaN(TotalAdvance) || double.IsInfinity(TotalAdvance))
                throw new ArgumentException("The pending whitespace advance must be finite.");
        }

        public IEnumerator<PendingSpaceFragment> GetEnumerator() {
            for (int index = _head; index < _items.Count; index++) yield return _items[index];
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
