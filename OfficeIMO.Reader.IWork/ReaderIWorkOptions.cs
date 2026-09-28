using OfficeIMO.IWork;

namespace OfficeIMO.Reader.IWork;

/// <summary>Controls the bounded Reader projection of Pages, Numbers, and Keynote sources.</summary>
public sealed class ReaderIWorkOptions {
    /// <summary>Gets or sets the source package and semantic limits.</summary>
    public IWorkReadOptions? ReadOptions { get; set; }

    /// <summary>Gets or sets the maximum number of columns materialized in each Reader table. Default: 256.</summary>
    public int MaximumTableColumns { get; set; } = 256;

    /// <summary>Gets or sets whether embedded image bytes are copied into the rich result. Default: false.</summary>
    public bool IncludeImagePayloads { get; set; }

    internal ReaderIWorkOptions Clone() {
        if (MaximumTableColumns < 1 || MaximumTableColumns > 16_384) {
            throw new ArgumentOutOfRangeException(nameof(MaximumTableColumns));
        }
        return new ReaderIWorkOptions {
            ReadOptions = ReadOptions?.Clone(),
            MaximumTableColumns = MaximumTableColumns,
            IncludeImagePayloads = IncludeImagePayloads
        };
    }
}
