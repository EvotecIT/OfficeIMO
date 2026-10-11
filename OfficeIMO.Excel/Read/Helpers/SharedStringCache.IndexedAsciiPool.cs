namespace OfficeIMO.Excel {
    internal sealed partial class SharedStringCache {
        /// <summary>
        /// Reuses one private ASCII index after its workbook closes. Entry storage has
        /// no text references, and retained capacity is bounded to 2 MiB of payload.
        /// Other sizes and simultaneous readers keep independently allocated lists.
        /// </summary>
        private static class IndexedAsciiEntryPool {
            private const int MinimumRetainedCapacity = 8192;
            private const int MaximumRetainedCapacity = 262144;
            private static readonly object Sync = new object();
            private static List<AsciiTextEntry>? _retained;

            internal static List<AsciiTextEntry> Rent(int minimumCapacity) {
                if (minimumCapacity >= MinimumRetainedCapacity
                    && minimumCapacity <= MaximumRetainedCapacity) {
                    lock (Sync) {
                        List<AsciiTextEntry>? retained = _retained;
                        if (retained != null && retained.Capacity >= minimumCapacity
                            && retained.Capacity <= (long)minimumCapacity * 2) {
                            _retained = null;
                            return retained;
                        }
                    }
                }

                // A cold read preserves the scanner's exact bounded capacity.
                return new List<AsciiTextEntry>(minimumCapacity);
            }

            internal static void Return(List<AsciiTextEntry> entries, bool retain = true) {
                // Reset the actual count. Every visible entry is overwritten by the
                // next scan; offset/length values contain no workbook text references.
                entries.Clear();
                if (!retain || entries.Capacity < MinimumRetainedCapacity
                    || entries.Capacity > MaximumRetainedCapacity) return;

                lock (Sync) {
                    // One slot follows the most recently completed eligible read.
                    _retained = entries;
                }
            }
        }
    }
}
