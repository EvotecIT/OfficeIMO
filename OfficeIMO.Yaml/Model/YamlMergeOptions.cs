namespace OfficeIMO.Yaml;

/// <summary>
/// How <see cref="YamlDocument.Merge"/> and <see cref="YamlMapping.Merge"/> combine an overlay sequence with the base
/// sequence at the same path: <see cref="Replace"/> (the default), <see cref="Append"/> or <see cref="MergeByKey"/>.
/// </summary>
public sealed class YamlSequenceMerge {
    private YamlSequenceMerge(YamlSequenceMergeMode mode, string? key) {
        Mode = mode;
        Key = key;
    }

    /// <summary>
    /// The overlay's items replace the base items, as Helm and <c>yq</c>'s <c>*</c> operator do. Items whose values
    /// already match keep their text and comments; only the items that differ change.
    /// </summary>
    public static YamlSequenceMerge Replace { get; } = new YamlSequenceMerge(YamlSequenceMergeMode.Replace, null);

    /// <summary>The overlay's items are appended after the base items.</summary>
    public static YamlSequenceMerge Append { get; } = new YamlSequenceMerge(YamlSequenceMergeMode.Append, null);

    /// <summary>
    /// Mapping items are matched by the value of a key, for example containers by <c>name</c>: an overlay item merges
    /// into the first base item with the same key value, and items without a match are appended.
    /// </summary>
    /// <param name="keyName">The key whose scalar value identifies an item.</param>
    /// <exception cref="ArgumentException"><paramref name="keyName"/> is null or empty.</exception>
    public static YamlSequenceMerge MergeByKey(string keyName) {
        if (string.IsNullOrEmpty(keyName)) throw new ArgumentException("A key name is required to match sequence items.", nameof(keyName));
        return new YamlSequenceMerge(YamlSequenceMergeMode.MergeByKey, keyName);
    }

    internal YamlSequenceMergeMode Mode { get; }

    internal string? Key { get; }
}

internal enum YamlSequenceMergeMode : byte {
    Replace,
    Append,
    MergeByKey
}

/// <summary>Settings for <see cref="YamlDocument.Merge"/> and <see cref="YamlMapping.Merge"/>.</summary>
public sealed class YamlMergeOptions {
    private YamlSequenceMerge _sequences = YamlSequenceMerge.Replace;

    internal static YamlMergeOptions Default { get; } = new YamlMergeOptions();

    /// <summary>Gets or sets how sequences at the same path combine. Defaults to <see cref="YamlSequenceMerge.Replace"/>.</summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public YamlSequenceMerge Sequences {
        get => _sequences;
        set => _sequences = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// Gets or sets whether an explicit null in the overlay removes the base key, Helm's convention for deleting a chart
    /// default. Off by default, so a null is an ordinary value that replaces. A null for a key the base does not have adds
    /// nothing when this is on.
    /// </summary>
    public bool NullRemovesKey { get; set; }
}
