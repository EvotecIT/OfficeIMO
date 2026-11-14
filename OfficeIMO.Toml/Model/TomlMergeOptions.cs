namespace OfficeIMO.Toml;

/// <summary>
/// How <see cref="TomlDocument.Merge"/> combines an overlay array with the base array at the same path:
/// <see cref="Replace"/> (the default), <see cref="Append"/> or <see cref="MergeByKey"/>. It applies to array values
/// (<c>[1, 2]</c>, arrays of inline tables) and to arrays of tables (<c>[[package]]</c>) alike, at every depth.
/// </summary>
public sealed class TomlArrayMerge {
    private TomlArrayMerge(TomlArrayMergeMode mode, string? key) {
        Mode = mode;
        Key = key;
    }

    /// <summary>
    /// The overlay's items replace the base items. Items whose values already match keep their text and comments; only the
    /// items that differ change.
    /// </summary>
    public static TomlArrayMerge Replace { get; } = new TomlArrayMerge(TomlArrayMergeMode.Replace, null);

    /// <summary>The overlay's items are appended after the base items.</summary>
    public static TomlArrayMerge Append { get; } = new TomlArrayMerge(TomlArrayMergeMode.Append, null);

    /// <summary>
    /// Table items are matched by the value of a key, for example the <c>[[package]]</c> tables of a lock file by
    /// <c>name</c>: an overlay table merges into the first base table with an equal value under that key that no earlier
    /// overlay table matched, so repeated values pair up in order, and tables without a match are appended. An array is
    /// matched this way only when it has items and every overlay item is a table with a value under the key; any other
    /// array, such as <c>features = ["std"]</c> or an empty <c>[]</c>, is replaced as with <see cref="Replace"/>.
    /// </summary>
    /// <param name="keyName">The key whose value identifies a table item.</param>
    /// <exception cref="ArgumentException"><paramref name="keyName"/> is null or empty.</exception>
    public static TomlArrayMerge MergeByKey(string keyName) {
        if (string.IsNullOrEmpty(keyName)) throw new ArgumentException("A key name is required to match array items.", nameof(keyName));
        return new TomlArrayMerge(TomlArrayMergeMode.MergeByKey, keyName);
    }

    internal TomlArrayMergeMode Mode { get; }

    internal string? Key { get; }
}

internal enum TomlArrayMergeMode : byte {
    Replace,
    Append,
    MergeByKey
}

/// <summary>Settings for <see cref="TomlDocument.Merge"/>.</summary>
public sealed class TomlMergeOptions {
    private TomlArrayMerge _arrays = TomlArrayMerge.Replace;

    internal static TomlMergeOptions Default { get; } = new TomlMergeOptions();

    /// <summary>
    /// Gets or sets how arrays and arrays of tables at the same path combine. Defaults to <see cref="TomlArrayMerge.Replace"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public TomlArrayMerge Arrays {
        get => _arrays;
        set => _arrays = value ?? throw new ArgumentNullException(nameof(value));
    }
}
