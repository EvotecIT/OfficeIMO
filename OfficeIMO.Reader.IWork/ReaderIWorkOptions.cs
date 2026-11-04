using OfficeIMO.IWork;

namespace OfficeIMO.Reader.IWork;

/// <summary>Controls the bounded Reader projection of Pages, Numbers, and Keynote sources.</summary>
public sealed class ReaderIWorkOptions {
    /// <summary>Gets or sets source package and semantic limits, also applied independently to emitted Reader image uses and comment payloads. Repeated uses count separately; comment text, author and native entry paths are charged before formatting.</summary>
    public IWorkReadOptions? ReadOptions { get; set; }

    /// <summary>Gets or sets the maximum number of columns materialized in each Reader table. Default: 256.</summary>
    public int MaxTableColumns { get; set; } = 256;

    /// <summary>Gets or sets the source-wide maximum number of dense table cells materialized in Reader output. Default: 1,000,000.</summary>
    public int MaxProjectedTableCells { get; set; } = 1_000_000;

    /// <summary>Gets or sets whether embedded image bytes are copied into the rich result. Default: false.</summary>
    public bool IncludeImagePayloads { get; set; }

    internal ReaderIWorkOptions Clone() {
        if (MaxTableColumns < 1 || MaxTableColumns > 16_384) {
            throw new ArgumentOutOfRangeException(nameof(MaxTableColumns));
        }
        if (MaxProjectedTableCells < 1) {
            throw new ArgumentOutOfRangeException(nameof(MaxProjectedTableCells));
        }
        return new ReaderIWorkOptions {
            ReadOptions = ReadOptions?.Clone(),
            MaxTableColumns = MaxTableColumns,
            MaxProjectedTableCells = MaxProjectedTableCells,
            IncludeImagePayloads = IncludeImagePayloads
        };
    }
}
