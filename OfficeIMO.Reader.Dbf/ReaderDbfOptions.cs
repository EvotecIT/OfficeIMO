using DBAClientX.Dbf;

namespace OfficeIMO.Reader.Dbf {
    /// <summary>Configures the text/table projection over the DbaClientX-owned typed DBF reader.</summary>
    public sealed class ReaderDbfOptions {
        /// <summary>Native codec limits, language driver override, and deleted-record policy. Caller streams always remain open.</summary>
        public DbfReadOptions ReadOptions { get; set; } = new();
        /// <summary>Maximum rows per emitted chunk, further constrained by ReaderOptions.MaxTableRows. Default: 200.</summary>
        public int ChunkRows { get; set; } = 200;
        /// <summary>Include a Markdown table projection. Default: true.</summary>
        public bool IncludeMarkdown { get; set; } = true;
        /// <summary>Allow path reads to open the same-stem DBT/FPT sidecar. Default: false. Stream and nested reads never resolve filesystem sidecars.</summary>
        public bool AllowMemoSidecarReads { get; set; }

        internal ReaderDbfOptions Copy() {
            if (ReadOptions == null) throw new ArgumentException("DBF read options are required.", nameof(ReadOptions));
            if (ChunkRows < 1 || ChunkRows > 10_000) throw new ArgumentOutOfRangeException(nameof(ChunkRows));
            return new ReaderDbfOptions {
                ChunkRows = ChunkRows, IncludeMarkdown = IncludeMarkdown, AllowMemoSidecarReads = AllowMemoSidecarReads,
                ReadOptions = new DbfReadOptions {
                    MaxInputBytes = ReadOptions.MaxInputBytes, MaxMemoFileBytes = ReadOptions.MaxMemoFileBytes,
                    MaxMemoBytes = ReadOptions.MaxMemoBytes, MaxTotalMemoBytes = ReadOptions.MaxTotalMemoBytes,
                    MaxRecords = ReadOptions.MaxRecords, MaxFields = ReadOptions.MaxFields,
                    IncludeDeletedRecords = ReadOptions.IncludeDeletedRecords,
                    Encoding = ReadOptions.Encoding == null ? null : (System.Text.Encoding)ReadOptions.Encoding.Clone(),
                    LeaveOpen = true
                }
            };
        }
    }
}
