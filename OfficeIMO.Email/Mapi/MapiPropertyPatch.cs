namespace OfficeIMO.Email;

/// <summary>Kind of change in a typed MAPI property patch.</summary>
public enum MapiPropertyPatchOperation {
    /// <summary>Set or replace one canonical property.</summary>
    Set = 0,
    /// <summary>Remove every property with one canonical identity.</summary>
    Remove = 1
}

/// <summary>One immutable change in a typed MAPI property patch.</summary>
public sealed class MapiPropertyPatchChange {
    internal MapiPropertyPatchChange(MapiPropertyPatchOperation operation, MapiPropertyKey key,
        object? value, MapiPropertyType? wireType, uint? flags) {
        Operation = operation;
        Key = key;
        Value = value;
        WireType = wireType;
        Flags = flags;
    }
    /// <summary>Operation to apply.</summary>
    public MapiPropertyPatchOperation Operation { get; }
    /// <summary>Canonical property identity.</summary>
    public MapiPropertyKey Key { get; }
    /// <summary>Replacement value, or null for removal.</summary>
    public object? Value { get; }
    /// <summary>Selected wire type for a set operation.</summary>
    public MapiPropertyType? WireType { get; }
    /// <summary>Optional MAPI property flags.</summary>
    public uint? Flags { get; }
}

/// <summary>
/// Ordered typed changes applied after normal semantic projection when an <see cref="EmailDocument"/> is written.
/// This lets advanced callers override or remove exact MAPI properties without replacing unrelated raw state.
/// </summary>
public sealed class MapiPropertyPatch {
    private readonly List<MapiPropertyPatchChange> _changes = new List<MapiPropertyPatchChange>();

    /// <summary>Ordered immutable view of the staged changes.</summary>
    public IReadOnlyList<MapiPropertyPatchChange> Changes => _changes;
    /// <summary>Whether the patch has no changes.</summary>
    public bool IsEmpty => _changes.Count == 0;

    /// <summary>Stages a typed set using the key's preferred wire type.</summary>
    public MapiPropertyPatch Set<T>(MapiPropertyKey<T> key, T value, uint? flags = null) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        return Set(key, value, key.PreferredType, flags);
    }

    /// <summary>Stages a typed set using an explicitly accepted wire type.</summary>
    public MapiPropertyPatch Set<T>(MapiPropertyKey<T> key, T value,
        MapiPropertyType wireType, uint? flags = null) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (!key.Accepts(wireType)) throw new ArgumentException(
            string.Concat(wireType.ToString(), " is not accepted by ", key.CanonicalName, "."), nameof(wireType));
        _changes.Add(new MapiPropertyPatchChange(MapiPropertyPatchOperation.Set,
            key, value, wireType, flags));
        return this;
    }

    /// <summary>Stages removal of every raw value with one canonical identity.</summary>
    public MapiPropertyPatch Remove(MapiPropertyKey key) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        _changes.Add(new MapiPropertyPatchChange(MapiPropertyPatchOperation.Remove,
            key, null, null, null));
        return this;
    }

    /// <summary>Appends a snapshot of another patch's ordered changes.</summary>
    public MapiPropertyPatch Append(MapiPropertyPatch patch) {
        if (patch == null) throw new ArgumentNullException(nameof(patch));
        _changes.AddRange(patch._changes);
        return this;
    }

    /// <summary>Applies the patch directly to a mutable raw property bag.</summary>
    public void Apply(MapiPropertyBag bag) {
        if (bag == null) throw new ArgumentNullException(nameof(bag));
        foreach (MapiPropertyPatchChange change in _changes) {
            if (change.Operation == MapiPropertyPatchOperation.Remove) bag.Remove(change.Key);
            else bag.SetValue(change.Key, change.Value!, change.WireType!.Value, change.Flags);
        }
    }

    internal void Apply(MsgPropertyBuilder builder) {
        foreach (MapiPropertyPatchChange change in _changes) {
            if (change.Operation == MapiPropertyPatchOperation.Remove) builder.Remove(change.Key);
            else builder.Set(change.Key, change.WireType!.Value, change.Value, change.Flags);
        }
    }
}
