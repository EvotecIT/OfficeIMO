using System.Runtime.CompilerServices;

namespace OfficeIMO.Yaml;

/// <summary>Compares objects by reference, for anchors and cycle detection.</summary>
internal sealed class YamlReferenceComparer : IEqualityComparer<object> {
    public static YamlReferenceComparer Instance { get; } = new YamlReferenceComparer();

    public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);

    public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
}

/// <summary>Per-call state of one deserialization: anchored instances and the path for error messages.</summary>
internal sealed class YamlReadState {
    private struct PathStep {
        public string? Key;
        public int Index;
    }

    [ThreadStatic] private static YamlReadState? t_spare;

    private PathStep[] _path = new PathStep[8];
    private int _depth;
    private Dictionary<string, object?>? _references;
    private Dictionary<YamlRecording, object?>? _shared;

    public YamlReadState(YamlSerializerOptions options, bool untypedRoot = false) {
        Options = options;
        Reset(options, untypedRoot);
    }

    /// <summary>
    /// Takes this thread's returned state, or creates one. A stream of many small documents then reads each one without
    /// a state of its own; pair it with <see cref="Return"/> once the value is read.
    /// </summary>
    public static YamlReadState Rent(YamlSerializerOptions options, bool untypedRoot) {
        YamlReadState? spare = t_spare;
        if (spare is null) return new YamlReadState(options, untypedRoot);
        t_spare = null;
        spare.Reset(options, untypedRoot);
        return spare;
    }

    /// <summary>Clears the state and keeps it for this thread's next read. The caller drops its reference.</summary>
    public void Return() {
        Array.Clear(_path, 0, _path.Length);
        _depth = 0;
        _references = null;
        _shared = null;
        if (_sizes is not null) Array.Clear(_sizes, 0, _sizes.Length);
        PendingAnchor = null;
        Options = null!;
        t_spare = this;
    }

    private void Reset(YamlSerializerOptions options, bool untypedRoot) {
        Options = options;
        Preserve = options.ReferenceHandling == YamlReferenceHandling.Preserve;
        SharesAliases = !Preserve && !options.CopyAliases;
        SharedTree = SharesAliases && untypedRoot;
    }

    public YamlSerializerOptions Options { get; private set; }

    public bool Preserve { get; private set; }

    /// <summary>Whether an alias read as <see cref="object"/> returns the value read for its anchor; see <see cref="YamlSerializerOptions.CopyAliases"/>.</summary>
    public bool SharesAliases { get; private set; }

    /// <summary>
    /// Whether the whole value is read as <see cref="object"/> with shared aliases, so every alias to a collection
    /// returns its shared value and nothing replays it: its events need not be kept.
    /// </summary>
    public bool SharedTree { get; private set; }

    /// <summary>Keeps the value read for an anchored node, for the aliases to it read later.</summary>
    public void Share(YamlRecording anchor, object? value) => (_shared ??= new Dictionary<YamlRecording, object?>())[anchor] = value;

    public bool TryGetShared(YamlRecording anchor, out object? value) {
        value = null;
        return _shared is not null && _shared.TryGetValue(anchor, out value);
    }

    /// <summary>The anchor of the node being read, until the converter that creates its instance takes it.</summary>
    public string? PendingAnchor { get; set; }

    public void PushKey(string key) {
        if (_depth == _path.Length) Array.Resize(ref _path, _depth * 2);
        _path[_depth++] = new PathStep { Key = key };
    }

    public void PushIndex(int index) {
        if (_depth == _path.Length) Array.Resize(ref _path, _depth * 2);
        _path[_depth++] = new PathStep { Index = index };
    }

    public void SetIndex(int index) => _path[_depth - 1].Index = index;

    // The sizes of the last mapping and sequence read at each depth. Collections at one depth usually repeat a shape, so
    // the next one starts at that size instead of growing to it step by step.
    private int[]? _sizes;
    private const int SizeHintLimit = 256;

    /// <summary>The size to create the next collection at the current depth with: the last one's, up to a limit.</summary>
    public int SizeHint(bool mapping) {
        int slot = 2 * _depth + (mapping ? 0 : 1);
        return _sizes is { } sizes && slot < sizes.Length ? sizes[slot] : 0;
    }

    /// <summary>Records the size of a collection just read at the current depth.</summary>
    public void NoteSize(bool mapping, int count) {
        int slot = 2 * _depth + (mapping ? 0 : 1);
        if (_sizes is null || slot >= _sizes.Length) Array.Resize(ref _sizes, Math.Max(16, slot * 2 + 2));
        _sizes[slot] = Math.Min(count, SizeHintLimit);
    }

    public void Pop() => _depth--;

    /// <summary>Registers the instance created for the pending anchor, before its children are read, so cycles resolve.</summary>
    public void TakePendingAnchor(object instance) {
        if (PendingAnchor is null) return;
        (_references ??= new Dictionary<string, object?>(StringComparer.Ordinal))[PendingAnchor] = instance;
        PendingAnchor = null;
    }

    public void Register(string anchor, object? value) =>
        (_references ??= new Dictionary<string, object?>(StringComparer.Ordinal))[anchor] = value;

    public bool TryGetReference(string anchor, out object? value) {
        value = null;
        return _references is not null && _references.TryGetValue(anchor, out value);
    }

    /// <summary>Formats the path of the node being read, in the path grammar of <see cref="YamlNode.Select"/>.</summary>
    public string FormatPath() {
        if (_depth == 0) return string.Empty;
        var text = new StringBuilder();
        for (int i = 0; i < _depth; i++) {
            PathStep step = _path[i];
            if (step.Key is null) {
                text.Append('[').Append(step.Index.ToString(CultureInfo.InvariantCulture)).Append(']');
            } else if (IsSimpleName(step.Key)) {
                if (text.Length > 0) text.Append('.');
                text.Append(step.Key);
            } else {
                text.Append("[\"").Append(step.Key.Replace("\\", "\\\\").Replace("\"", "\\\"")).Append("\"]");
            }
        }
        return text.ToString();
    }

    private static bool IsSimpleName(string key) {
        if (key.Length == 0) return false;
        foreach (char c in key) {
            if (c is '.' or '[' or ']' or '"' or '$' || char.IsWhiteSpace(c) || char.IsControl(c)) return false;
        }
        return true;
    }
}

/// <summary>Per-call state of one serialization: anchors for shared references and the stack for cycle detection.</summary>
internal sealed class YamlWriteState {
    private const int ScannedAncestors = 8;
    // A struct element stores without the covariance check of an object array. Leaving a level does not clear its
    // slot; Return clears the slots used, so a kept state holds no references.
    private struct Ancestor {
        public object Value;
    }
    private Ancestor[] _ancestors = new Ancestor[8];
    private int _maxDepth;
    // Ancestors past the scanned ones also sit in an open-addressed identity table, with the slot each took. They
    // leave in reverse order of entering, so leaving clears the slot: every entry still present was placed earlier,
    // so no probe sequence of one runs through a slot taken later.
    private Ancestor[] _deepTable = Array.Empty<Ancestor>();
    private int[] _deepSlots = Array.Empty<int>();
    private int _depth;
    // See Enter: a write whose output is discarded on failure checks for cycles by one comparison per level.
    private bool _quick;
    private const int QuickDepthLimit = 256;
    private Dictionary<object, string>? _anchors;
    private HashSet<object>? _shared;
    private int _nextAnchor;

    // One state per thread is reused by the next call after a completed write; a call that throws, or a nested call
    // while one is in use, creates its own.
    [ThreadStatic] private static YamlWriteState? t_cached;

    private YamlWriteState(YamlSerializerOptions options) {
        Options = options;
        Preserve = options.ReferenceHandling == YamlReferenceHandling.Preserve;
    }

    /// <summary>Gets a state for one write with the options; pair it with <see cref="Return"/> when the write completes.</summary>
    /// <remarks>
    /// With <paramref name="quickCycleCheck"/>, set only when the output of a write that fails is discarded, cycles are
    /// suspected by one comparison per level (see <see cref="Enter"/>); the caller writes again with the exact check
    /// when <see cref="CycleSuspected"/> is set, so the result or error is the exact check's.
    /// </remarks>
    public static YamlWriteState Rent(YamlSerializerOptions options, bool quickCycleCheck = false) {
        YamlWriteState? state = t_cached;
        if (state is null) state = new YamlWriteState(options);
        else t_cached = null;
        return state.Start(options, quickCycleCheck);
    }

    /// <summary>Gets a state as <see cref="Rent"/> does, starting from <paramref name="spare"/> when there is one.</summary>
    public static YamlWriteState Reuse(YamlWriteState? spare, YamlSerializerOptions options, bool quickCycleCheck) =>
        spare is null ? Rent(options, quickCycleCheck) : spare.Start(options, quickCycleCheck);

    private YamlWriteState Start(YamlSerializerOptions options, bool quickCycleCheck) {
        Options = options;
        Preserve = options.ReferenceHandling == YamlReferenceHandling.Preserve;
        _quick = quickCycleCheck;
        CycleSuspected = false;
        return this;
    }

    /// <summary>Whether the quick check stopped the write at a possible cycle; the write must be repeated with the exact check.</summary>
    public bool CycleSuspected { get; private set; }

    /// <summary>Clears the state of a completed write and keeps it for the next write on this thread.</summary>
    public void Return() {
        if (Clear()) t_cached = this;
    }

    /// <summary>Clears the state of a write; true when it completed, so the state can be used again.</summary>
    public bool Clear() {
        // A completed write has left every object it entered; the deep table holds no references.
        Array.Clear(_ancestors, 0, Math.Min(_maxDepth, _ancestors.Length));
        _maxDepth = 0;
        _anchors = null;
        _shared = null;
        _nextAnchor = 0;
        DiscriminatorKey = null;
        DiscriminatorValue = null;
        Options = null!;
        return _depth == 0;
    }

    public YamlSerializerOptions Options { get; private set; }

    public bool Preserve { get; private set; }

    /// <summary>The discriminator key and value the next object mapping writes first, for property discriminators.</summary>
    public string? DiscriminatorKey { get; set; }

    public string? DiscriminatorValue { get; set; }

    /// <summary>Sets the objects reached more than once, found by walking the graph before writing.</summary>
    public void SetShared(HashSet<object> shared) => _shared = shared;

    /// <summary>
    /// Returns the alias to write for an object already written, or null; for a shared object reached the first time it
    /// assigns the anchor to write.
    /// </summary>
    public bool TryGetAlias(object value, out string? alias, out string? anchor) {
        alias = null;
        anchor = null;
        if (_shared is null || !_shared.Contains(value)) return false;
        _anchors ??= new Dictionary<object, string>(YamlReferenceComparer.Instance);
        if (_anchors.TryGetValue(value, out alias)) return true;
        anchor = "id" + (++_nextAnchor).ToString("000", CultureInfo.InvariantCulture);
        _anchors.Add(value, anchor);
        return false;
    }

    /// <summary>
    /// Starts writing an object or collection. Under Preserve it writes the alias of an object already written and returns
    /// false, or sets the anchor of a shared object; otherwise it fails on a cycle. Pair a true result with <see cref="End"/>.
    /// </summary>
    public bool Begin(YamlWriter writer, object value, Type type) {
        if (Preserve) {
            if (TryGetAlias(value, out string? alias, out string? anchor)) {
                writer.WriteAlias(alias!);
                return false;
            }
            if (anchor is not null) writer.WriteAnchor(anchor);
            return true;
        }
        Enter(value, type);
        return true;
    }

    public void End() {
        if (!Preserve) Exit();
    }

    /// <summary>
    /// Enters an object or collection, failing on a cycle. The first <see cref="ScannedAncestors"/> ancestors are
    /// compared one by one; deeper ones are also kept in an identity table, so deep nesting costs a lookup per level
    /// instead of a scan of the whole path.
    /// </summary>
    /// <remarks>
    /// The quick check compares each value only with the ancestor at the highest power of two below its depth. A cycle
    /// repeats its path with a fixed period, so a level of it matches that ancestor before the path is about four times
    /// as deep as the cycle's start and length; a match, or a path as deep as <see cref="QuickDepthLimit"/>, stops the
    /// write with <see cref="YamlCycleSuspectedException"/>.
    /// </remarks>
    public void Enter(object value, Type type) {
        int depth = _depth;
        Ancestor[] ancestors = _ancestors;
        if (_quick) {
            if (depth > 0 && (ReferenceEquals(ancestors[HighestPowerOfTwoBelow(depth)].Value, value) || depth >= QuickDepthLimit)) {
                CycleSuspected = true;
                throw new YamlCycleSuspectedException();
            }
            if (depth == ancestors.Length) {
                Array.Resize(ref _ancestors, depth * 2);
                ancestors = _ancestors;
            }
            ancestors[depth].Value = value;
            _depth = depth + 1;
            if (depth >= _maxDepth) _maxDepth = depth + 1;
            return;
        }
        int scanned = Math.Min(depth, ScannedAncestors);
        int i = 0;
        // Four at a time with one branch: a deep tree compares many ancestors per level.
        for (; i + 4 <= scanned; i += 4) {
            if (ReferenceEquals(ancestors[i].Value, value) | ReferenceEquals(ancestors[i + 1].Value, value)
                | ReferenceEquals(ancestors[i + 2].Value, value) | ReferenceEquals(ancestors[i + 3].Value, value)) throw YamlSerializationErrors.Cycle(type);
        }
        for (; i < scanned; i++) {
            if (ReferenceEquals(ancestors[i].Value, value)) throw YamlSerializationErrors.Cycle(type);
        }
        if (depth >= ScannedAncestors && !AddDeep(value, depth - ScannedAncestors)) throw YamlSerializationErrors.Cycle(type);
        if (depth == ancestors.Length) {
            Array.Resize(ref _ancestors, depth * 2);
            ancestors = _ancestors;
        }
        ancestors[depth].Value = value;
        _depth = depth + 1;
        if (depth >= _maxDepth) _maxDepth = depth + 1;
    }

    public void Exit() {
        if (_quick) {
            _depth--;
            return;
        }
        int depth = --_depth;
        if (depth >= ScannedAncestors) _deepTable[_deepSlots[depth - ScannedAncestors]].Value = null!;
    }

    /// <summary>The highest power of two below <paramref name="depth"/>, which is at least 1: 0 for 1, 1 for 2, 2 for 3 and 4.</summary>
    private static int HighestPowerOfTwoBelow(int depth) {
        int below = depth - 1;
        if (below == 0) return 0;
#if NET8_0_OR_GREATER
        return 1 << System.Numerics.BitOperations.Log2((uint)below);
#else
        int power = 1;
        while (power * 2 <= below) power *= 2;
        return power;
#endif
    }

    /// <summary>Adds a deep ancestor as the table's entry number <paramref name="count"/>; false when it is already there.</summary>
    private bool AddDeep(object value, int count) {
        if ((count + 1) * 2 > _deepTable.Length) GrowDeep(count);
        Ancestor[] table = _deepTable;
        int mask = table.Length - 1;
        int slot = RuntimeHelpers.GetHashCode(value) & mask;
        while (table[slot].Value is object present) {
            if (ReferenceEquals(present, value)) return false;
            slot = (slot + 1) & mask;
        }
        table[slot].Value = value;
        if (count == _deepSlots.Length) Array.Resize(ref _deepSlots, Math.Max(16, count * 2));
        _deepSlots[count] = slot;
        return true;
    }

    /// <summary>Doubles the table and places the present entries again, in the order they entered.</summary>
    private void GrowDeep(int count) {
        var table = new Ancestor[Math.Max(64, _deepTable.Length * 2)];
        int mask = table.Length - 1;
        for (int i = 0; i < count; i++) {
            object value = _ancestors[ScannedAncestors + i].Value;
            int slot = RuntimeHelpers.GetHashCode(value) & mask;
            while (table[slot].Value is not null) slot = (slot + 1) & mask;
            table[slot].Value = value;
            _deepSlots[i] = slot;
        }
        _deepTable = table;
    }
}

/// <summary>Stops a write whose quick cycle check found a possible cycle; the serializer writes again with the exact check.</summary>
internal sealed class YamlCycleSuspectedException : Exception {
}

/// <summary>Walks an object graph before writing with <see cref="YamlReferenceHandling.Preserve"/> to find shared references.</summary>
internal sealed class YamlReferenceCollector {
    private readonly HashSet<object> _seen = new HashSet<object>(YamlReferenceComparer.Instance);

    public YamlReferenceCollector(YamlSerializerOptions options) {
        Options = options;
    }

    public YamlSerializerOptions Options { get; }

    public HashSet<object> Shared { get; } = new HashSet<object>(YamlReferenceComparer.Instance);

    /// <summary>Returns true the first time an object is visited; a second visit marks it shared.</summary>
    public bool Visit(object value) {
        if (_seen.Add(value)) return true;
        Shared.Add(value);
        return false;
    }
}
