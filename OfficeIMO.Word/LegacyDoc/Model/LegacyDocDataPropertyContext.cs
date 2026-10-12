namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Owns one document's bounded Data-property traversal and reusable resolved records.</summary>
    internal sealed class LegacyDocDataPropertyContext {
        private const int MaximumCachedBytes = 8 * 1024 * 1024;
        private const int MaximumCachedRecords = 4096;
        private readonly Dictionary<int, CachedRecord> records = new();
        private int cachedBytes;
        private int remainingWorkBytes;

        internal LegacyDocDataPropertyContext(byte[] dataStream, LegacyDocImportOptions options) {
            options.Validate();
            DataStream = dataStream;
            MaximumChainLength = options.MaxDataPropertyChainLength;
            remainingWorkBytes = options.MaxParagraphPropertyWorkBytes;
        }

        internal byte[] DataStream { get; }
        internal int MaximumChainLength { get; }
        internal bool WorkLimitExceeded { get; private set; }

        internal void ConsumeWork(int bytes) {
            if (bytes < 0 || bytes > remainingWorkBytes) {
                WorkLimitExceeded = true;
                remainingWorkBytes = 0;
                throw new InvalidDataException("The document's paragraph properties exceed the configured traversal and expansion work limit.");
            }
            remainingWorkBytes -= bytes;
        }

        internal bool TryGetRecord(int offset, out byte[] properties, out int followedRecords) {
            if (records.TryGetValue(offset, out CachedRecord? record)) {
                properties = record.Properties;
                followedRecords = record.FollowedRecords;
                return true;
            }
            properties = Array.Empty<byte>();
            followedRecords = 0;
            return false;
        }

        internal void CacheRecord(int offset, byte[] properties, int prefixLength, int followedRecords) {
            int length = properties.Length - prefixLength;
            if (records.ContainsKey(offset) || records.Count >= MaximumCachedRecords ||
                length > Math.Min(MaximumCachedBytes - cachedBytes, remainingWorkBytes)) return;
            var copy = new byte[length];
            Buffer.BlockCopy(properties, prefixLength, copy, 0, length);
            records.Add(offset, new CachedRecord(copy, followedRecords));
            cachedBytes += length;
        }

        private sealed class CachedRecord {
            internal CachedRecord(byte[] properties, int followedRecords) {
                Properties = properties;
                FollowedRecords = followedRecords;
            }
            internal byte[] Properties { get; }
            internal int FollowedRecords { get; }
        }
    }
}
