namespace OfficeIMO.Excel {
    /// <summary>
    /// Keeps large arrays for reuse by the next read or save in a batch without holding them once the work stops.
    /// A returned array is held strongly until a full garbage collection finds it unused for a few seconds and weakly
    /// after that, so the collector frees it unless another operation takes it first.
    /// </summary>
    internal sealed class ExcelReleasingArrayStore<T> {
        // Operations in a batch reuse their arrays well within this time; an idle process gives them up.
        private static readonly long IdleTicks = System.Diagnostics.Stopwatch.Frequency * 5;
        private readonly object _sync = new object();
        private readonly Dictionary<int, List<Slot>> _slots = new Dictionary<int, List<Slot>>();
        private readonly int _maximumPerLength;
        private readonly long _maximumStrongElements;
        private long _strongElements;
        private bool _trimRegistered;

        internal ExcelReleasingArrayStore(int maximumPerLength, long maximumStrongElements) {
            _maximumPerLength = maximumPerLength;
            _maximumStrongElements = maximumStrongElements;
        }

        /// <summary>Takes a retained array of exactly <paramref name="length"/> elements, if one is still alive.</summary>
        internal T[]? TryTake(int length) {
            lock (_sync) {
                if (!_slots.TryGetValue(length, out List<Slot>? slots)) return null;
                for (int index = slots.Count - 1; index >= 0; index--) {
                    Slot slot = slots[index];
                    slots.RemoveAt(index);
                    if (slot.Strong != null) {
                        _strongElements -= length;
                        return slot.Strong;
                    }

                    if (slot.Weak!.TryGetTarget(out T[]? array)) return array;
                }

                return null;
            }
        }

        /// <summary>Retains <paramref name="array"/> for reuse; returns false when the store is full.</summary>
        internal bool TryAdd(T[] array) {
            lock (_sync) {
                if (!_slots.TryGetValue(array.Length, out List<Slot>? slots)) {
                    slots = new List<Slot>(_maximumPerLength);
                    _slots.Add(array.Length, slots);
                }

                slots.RemoveAll(slot => slot.Strong == null && !slot.Weak!.TryGetTarget(out _));
                if (slots.Count >= _maximumPerLength || _strongElements + array.Length > _maximumStrongElements) {
                    return false;
                }

                slots.Add(new Slot(array, GC.CollectionCount(2), System.Diagnostics.Stopwatch.GetTimestamp()));
                _strongElements += array.Length;
                if (!_trimRegistered) {
                    _trimRegistered = true;
                    FullCollectionTrim.Register(this);
                }

                return true;
            }
        }

        /// <summary>Weakens arrays unused since before the last full collection and for a few seconds; returns whether strongly held arrays remain.</summary>
        private bool Trim() {
            int fullCollections = GC.CollectionCount(2);
            long now = System.Diagnostics.Stopwatch.GetTimestamp();
            lock (_sync) {
                bool remaining = false;
                foreach (List<Slot> slots in _slots.Values) {
                    for (int index = slots.Count - 1; index >= 0; index--) {
                        Slot slot = slots[index];
                        if (slot.Strong == null) {
                            if (!slot.Weak!.TryGetTarget(out _)) slots.RemoveAt(index);
                        } else if (slot.FullCollectionsAtReturn < fullCollections && now - slot.ReturnedAt >= IdleTicks) {
                            _strongElements -= slot.Strong.Length;
                            slots[index] = new Slot(new WeakReference<T[]>(slot.Strong));
                        } else {
                            remaining = true;
                        }
                    }
                }

                if (!remaining) _trimRegistered = false;
                return remaining;
            }
        }

        private readonly struct Slot {
            internal Slot(T[] strong, int fullCollectionsAtReturn, long returnedAt) {
                Strong = strong;
                Weak = null;
                FullCollectionsAtReturn = fullCollectionsAtReturn;
                ReturnedAt = returnedAt;
            }

            internal Slot(WeakReference<T[]> weak) {
                Strong = null;
                Weak = weak;
                FullCollectionsAtReturn = 0;
                ReturnedAt = 0;
            }

            internal T[]? Strong { get; }
            internal WeakReference<T[]>? Weak { get; }
            internal int FullCollectionsAtReturn { get; }
            internal long ReturnedAt { get; }
        }

        /// <summary>
        /// An unreachable finalizable object: once it reaches generation 2 its finalizer runs after each full
        /// collection, which is when unused arrays are weakened.
        /// </summary>
        private sealed class FullCollectionTrim {
            private readonly ExcelReleasingArrayStore<T> _store;

            private FullCollectionTrim(ExcelReleasingArrayStore<T> store) => _store = store;

            internal static void Register(ExcelReleasingArrayStore<T> store) => _ = new FullCollectionTrim(store);

            ~FullCollectionTrim() {
                if (Environment.HasShutdownStarted) return;
                if (_store.Trim()) GC.ReRegisterForFinalize(this);
            }
        }
    }
}
