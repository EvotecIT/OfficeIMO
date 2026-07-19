namespace OfficeIMO.Email;

/// <summary>Mutable Outlook category names with case-insensitive convenience operations.</summary>
public sealed class OutlookCategoryCollection : IList<string> {
    private readonly List<string> _items = new List<string>();

    /// <summary>Adds a non-empty category unless an equivalent name already exists.</summary>
    /// <returns><see langword="true"/> when the category was added.</returns>
    public bool AddIfMissing(string category) {
        string normalized = Validate(category);
        if (Contains(normalized, StringComparer.OrdinalIgnoreCase)) return false;
        _items.Add(normalized);
        return true;
    }

    /// <summary>Replaces the collection, optionally removing case-insensitive duplicates.</summary>
    public void ReplaceWith(IEnumerable<string> categories, bool removeDuplicates = true) {
        if (categories == null) throw new ArgumentNullException(nameof(categories));
        string[] replacement = categories.Select(Validate).ToArray();
        _items.Clear();
        foreach (string category in replacement) {
            if (!removeDuplicates || !Contains(category, StringComparer.OrdinalIgnoreCase)) _items.Add(category);
        }
    }

    /// <summary>Removes every category matching the supplied name case-insensitively.</summary>
    /// <returns>The number of removed category entries.</returns>
    public int RemoveAll(string category) {
        string normalized = Validate(category);
        int removed = 0;
        for (int index = _items.Count - 1; index >= 0; index--) {
            if (!string.Equals(_items[index], normalized, StringComparison.OrdinalIgnoreCase)) continue;
            _items.RemoveAt(index);
            removed++;
        }
        return removed;
    }

    /// <summary>Returns whether the collection contains a category using the requested comparison.</summary>
    public bool Contains(string category, StringComparer comparer) {
        if (comparer == null) throw new ArgumentNullException(nameof(comparer));
        string normalized = Validate(category);
        return _items.Any(existing => comparer.Equals(existing, normalized));
    }

    /// <inheritdoc />
    public IEnumerator<string> GetEnumerator() => _items.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

    /// <inheritdoc />
    public void Add(string item) => _items.Add(Validate(item));

    /// <inheritdoc />
    public void Clear() => _items.Clear();

    /// <inheritdoc />
    public bool Contains(string item) => item != null && _items.Contains(item);

    /// <inheritdoc />
    public void CopyTo(string[] array, int arrayIndex) => _items.CopyTo(array, arrayIndex);

    /// <inheritdoc />
    public bool Remove(string item) => item != null && _items.Remove(item);

    /// <inheritdoc />
    public int Count => _items.Count;

    /// <inheritdoc />
    public bool IsReadOnly => false;

    /// <inheritdoc />
    public int IndexOf(string item) => item == null ? -1 : _items.IndexOf(item);

    /// <inheritdoc />
    public void Insert(int index, string item) => _items.Insert(index, Validate(item));

    /// <inheritdoc />
    public void RemoveAt(int index) => _items.RemoveAt(index);

    /// <inheritdoc />
    public string this[int index] {
        get => _items[index];
        set => _items[index] = Validate(value);
    }

    private static string Validate(string category) {
        if (category == null) throw new ArgumentNullException(nameof(category));
        string normalized = category.Trim();
        if (normalized.Length == 0) throw new ArgumentException("A category name cannot be empty.", nameof(category));
        return normalized;
    }
}
