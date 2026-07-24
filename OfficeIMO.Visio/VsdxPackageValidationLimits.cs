namespace OfficeIMO.Visio {
    /// <summary>Resource limits applied while validating or repairing VSDX packages.</summary>
    public sealed class VsdxPackageValidationLimits {
        /// <summary>Maximum number of ZIP entries. Default: 4096.</summary>
        public int MaxEntries { get; set; } = 4096;

        /// <summary>Maximum uncompressed size of one entry. Default: 64 MiB.</summary>
        public long MaxEntryBytes { get; set; } = 64L * 1024L * 1024L;

        /// <summary>Maximum aggregate uncompressed size. Default: 256 MiB.</summary>
        public long MaxTotalBytes { get; set; } = 256L * 1024L * 1024L;

        /// <summary>Maximum uncompressed-to-compressed ratio for a non-empty entry. Default: 200.</summary>
        public double MaxCompressionRatio { get; set; } = 200d;

        internal Snapshot CreateSnapshot() {
            if (MaxEntries <= 0) throw new ArgumentOutOfRangeException(nameof(MaxEntries));
            if (MaxEntryBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxEntryBytes));
            if (MaxTotalBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalBytes));
            if (MaxCompressionRatio <= 0 || double.IsNaN(MaxCompressionRatio) || double.IsInfinity(MaxCompressionRatio)) {
                throw new ArgumentOutOfRangeException(nameof(MaxCompressionRatio));
            }
            return new Snapshot(MaxEntries, MaxEntryBytes, MaxTotalBytes, MaxCompressionRatio);
        }

        internal readonly struct Snapshot {
            internal Snapshot(int maxEntries, long maxEntryBytes, long maxTotalBytes, double maxCompressionRatio) {
                MaxEntries = maxEntries;
                MaxEntryBytes = maxEntryBytes;
                MaxTotalBytes = maxTotalBytes;
                MaxCompressionRatio = maxCompressionRatio;
            }

            internal int MaxEntries { get; }
            internal long MaxEntryBytes { get; }
            internal long MaxTotalBytes { get; }
            internal double MaxCompressionRatio { get; }
        }
    }
}
