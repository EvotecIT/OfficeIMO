using OfficeIMO.IWork;

namespace OfficeIMO.Reader.IWork;

/// <summary>Controls the bounded Reader projection of Pages, Numbers, and Keynote sources.</summary>
public sealed class ReaderIWorkOptions {
    /// <summary>Gets or sets the source package, semantic, and emitted Reader image-use limits. Repeated Reader image uses count separately; payload bytes are charged before copying.</summary>
    public IWorkReadOptions? ReadOptions { get; set; }

    /// <summary>Gets or sets the maximum number of columns materialized in each Reader table. Default: 256.</summary>
    public int MaximumTableColumns { get; set; } = 256;

    /// <summary>Gets or sets the source-wide maximum number of dense table cells materialized in Reader output. Default: 1,000,000.</summary>
    public int MaximumProjectedTableCells { get; set; } = 1_000_000;

    /// <summary>Gets or sets whether embedded image bytes are copied into the rich result. Default: false.</summary>
    public bool IncludeImagePayloads { get; set; }

    internal ReaderIWorkOptions Clone() {
        if (MaximumTableColumns < 1 || MaximumTableColumns > 16_384) {
            throw new ArgumentOutOfRangeException(nameof(MaximumTableColumns));
        }
        if (MaximumProjectedTableCells < 1) {
            throw new ArgumentOutOfRangeException(nameof(MaximumProjectedTableCells));
        }
        return new ReaderIWorkOptions {
            ReadOptions = ReadOptions?.Clone(),
            MaximumTableColumns = MaximumTableColumns,
            MaximumProjectedTableCells = MaximumProjectedTableCells,
            IncludeImagePayloads = IncludeImagePayloads
        };
    }
}
