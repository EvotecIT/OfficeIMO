namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static ulong GetComparisonFingerprint<T>(T value) where T : notnull {
            if (value is string text) {
                return GetOrdinalTextFingerprint(text);
            }

            if (value is IComparisonFingerprint fingerprint) {
                return fingerprint.ComparisonFingerprint;
            }

            throw new InvalidOperationException(
                "Structured comparison values must expose an independent comparison fingerprint.");
        }

        private static ulong GetOrdinalTextFingerprint(string value) {
            const ulong offsetBasis = 14695981039346656037UL;
            const ulong prime = 1099511628211UL;
            ulong fingerprint = offsetBasis;
            unchecked {
                for (int index = 0; index < value.Length; index++) {
                    char character = value[index];
                    fingerprint ^= (byte)character;
                    fingerprint *= prime;
                    fingerprint ^= (byte)(character >> 8);
                    fingerprint *= prime;
                }

                fingerprint ^= (uint)value.Length;
                fingerprint *= prime;
            }

            return fingerprint;
        }

        private static ulong CombineComparisonFingerprints(ulong first, ulong second) {
            unchecked {
                return (first * 11400714819323198485UL) ^ second;
            }
        }

        private sealed class ComparisonWorkBudget {
            private long _remainingWorkUnits;

            internal ComparisonWorkBudget(long maximumWorkUnits) {
                _remainingWorkUnits = Math.Max(0, maximumWorkUnits);
            }

            internal bool IsExhausted => _remainingWorkUnits <= 0;

            internal bool TryConsume(long workUnits) {
                if (workUnits <= 0) {
                    return true;
                }

                if (workUnits > _remainingWorkUnits) {
                    _remainingWorkUnits = 0;
                    return false;
                }

                _remainingWorkUnits -= workUnits;
                return true;
            }
        }

        internal interface IComparisonFingerprint {
            ulong ComparisonFingerprint { get; }
        }
    }
}
