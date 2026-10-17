namespace OfficeIMO.Archives {

    /// <summary>Finite cumulative admission allowances for a caller-owned tree of archive operations.</summary>
    public sealed class ArchiveOperationLimits {
        /// <summary>Aggregate compressed input admitted by all containers, including nested inputs.</summary>
        public long MaxInputBytes { get; set; } = 1024L * 1024 * 1024;
        /// <summary>Aggregate physical entries, including MSI File rows and embedded cabinet members.</summary>
        public long MaxEntryCount { get; set; } = 200000;
        /// <summary>Aggregate attempted decoded expansion, including repeated reads and solid-folder siblings.</summary>
        public long MaxExpandedBytes { get; set; } = 2L * 1024 * 1024 * 1024;
        /// <summary>Aggregate selected payload bytes admitted for copying or returning. Charges are not refunded on disposal or failure.</summary>
        public long MaxSelectedBytes { get; set; } = 1024L * 1024 * 1024;
        /// <summary>Aggregate processing units: CAB frame charges, and compressed plus twice expanded bytes for ZIP reads.</summary>
        public long MaxDecompressionWorkBytes { get; set; } = 4L * 1024 * 1024 * 1024;
        /// <summary>Maximum nested container depth, with a root container at zero.</summary>
        public int MaxDepth { get; set; } = 8;
    }
}
