using System.Collections;

namespace OfficeIMO.Yaml;

/// <summary>
/// A depth-first, document-order walk over a node and everything below it, returned by <see cref="YamlNode.Descendants"/>.
/// <c>foreach</c> over it allocates nothing per node: the pending nodes live on one pooled stack that the enumerator
/// returns when it is disposed.
/// </summary>
public readonly struct YamlNodeWalk : IEnumerable<YamlNode> {
    private readonly YamlNode _root;

    internal YamlNodeWalk(YamlNode root) => _root = root;

    /// <summary>Returns the enumerator.</summary>
    public Enumerator GetEnumerator() => new Enumerator(_root);

    IEnumerator<YamlNode> IEnumerable<YamlNode>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Enumerates the walk. Dispose it (as <c>foreach</c> does) to return its pooled stack.</summary>
    public struct Enumerator : IEnumerator<YamlNode> {
        private readonly YamlNode _root;
        private YamlNode[]? _stack;
        private int _count;
        private YamlNode? _current;

        internal Enumerator(YamlNode root) {
            _root = root;
            _stack = null;
            _count = -1;
            _current = null;
        }

        /// <inheritdoc />
        public YamlNode Current => _current!;

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public bool MoveNext() {
            if (_count < 0) {
                if (_root is null) {
                    _count = 0;
                    return false;
                }
                _stack = ArrayPool<YamlNode>.Shared.Rent(16);
                _stack[0] = _root;
                _count = 1;
            }
            if (_count == 0) {
                _current = null;
                return false;
            }
            YamlNode node = _stack![--_count];
            _stack[_count] = null!;
            _current = node;
            if (node is YamlMapping mapping) {
                Reserve(mapping._count * 2);
                for (int i = mapping._count - 1; i >= 0; i--) {
                    _stack[_count++] = mapping.Entries[i].Value;
                    _stack[_count++] = mapping.Entries[i].Key;
                }
            } else if (node is YamlSequence sequence) {
                Reserve(sequence._count);
                for (int i = sequence._count - 1; i >= 0; i--) _stack[_count++] = sequence.Items[i];
            }
            return true;
        }

        /// <inheritdoc />
        public void Reset() {
            Dispose();
            _count = -1;
        }

        /// <inheritdoc />
        public void Dispose() {
            if (_stack is null) return;
            Array.Clear(_stack, 0, Math.Max(_count, 0));
            ArrayPool<YamlNode>.Shared.Return(_stack);
            _stack = null;
            _count = 0;
        }

        private void Reserve(int extra) {
            if (_count + extra <= _stack!.Length) return;
            YamlNode[] larger = ArrayPool<YamlNode>.Shared.Rent(Math.Max(_stack.Length * 2, _count + extra));
            Array.Copy(_stack, larger, _count);
            Array.Clear(_stack, 0, _count);
            ArrayPool<YamlNode>.Shared.Return(_stack);
            _stack = larger;
        }
    }
}
